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

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer)
    'Guardia caos

    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

    'Npclist(NpcIndex).CanAttack = 0
    Dim daño As Integer

    If Hechizos(Spell).SubeHP = 1 Then

        daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + daño

        If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
    
        Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(UserIndex)
        Call SubirSkill(UserIndex, Resistencia)

    ElseIf Hechizos(Spell).SubeHP = 2 Then
        
        daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)

        End If
        
        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)

        End If
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
                
        If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
          
        End If
        
        If UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) > 0 Then

            Dim DefensaMagica As Long

            Dim Absorcion     As Long

            DefensaMagica = UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 4
            Absorcion = daño / 100 * DefensaMagica
            daño = daño - Absorcion

        End If

        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - daño
        
        Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(UserIndex)
        Call SubirSkill(UserIndex, Resistencia)
        
        'Muere
        If UserList(UserIndex).Stats.MinHp < 1 Then
            UserList(UserIndex).Stats.MinHp = 0
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
    
    ElseIf Hechizos(Spell).Paraliza = 1 Then

        If UserList(UserIndex).flags.Paralizado = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

            UserList(UserIndex).flags.Paralizado = 1
            UserList(UserIndex).Counters.Paralisis = Hechizos(Spell).Duration / 2
          
            Call WriteParalizeOK(UserIndex)
            Call WritePosUpdate(UserIndex)
        End If

    ElseIf Hechizos(Spell).incinera = 1 Then
        Debug.Print "incinerar"

        If UserList(UserIndex).flags.Incinerado = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

            If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

            End If

            UserList(UserIndex).flags.Incinerado = 1
            Call WriteConsoleMsg(UserIndex, "Has sido incinerado por " & Npclist(NpcIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End If

End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
    'solo hechizos ofensivos!

    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
    Npclist(NpcIndex).CanAttack = 0

    Dim daño As Integer

    If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        
        Npclist(TargetNPC).Stats.MinHp = Npclist(TargetNPC).Stats.MinHp - daño
        
        'Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, daño)
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHp < 1 Then
            Npclist(TargetNPC).Stats.MinHp = 0
            ' If Npclist(NpcIndex).MaestroUser > 0 Then
            '  Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            '  Else
            Call MuereNpc(TargetNPC, 0)

            '  End If
        End If
    
    End If
    
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

    On Error GoTo Errhandler
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function

        End If

    Next

    Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal slot As Integer)

    Dim hIndex As Integer

    Dim j      As Integer

    hIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex

    If Not TieneHechizo(hIndex, UserIndex) Then

        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS

            If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
        Next j
        
        If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
        Else
            UserList(UserIndex).Stats.UserHechizos(j) = hIndex
            Call UpdateUserHechizos(False, UserIndex, CByte(j))
            'Quitamos del inv el item
            Call QuitarUserInvItem(UserIndex, CByte(slot), 1)

        End If

    Else
        Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Byte, ByVal UserIndex As Integer)

    On Error Resume Next

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.CharIndex, vbCyan))
    Exit Sub

End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer, Optional ByVal slot As Integer = 0) As Boolean

    If UserList(UserIndex).flags.Muerto = 0 Then

        Dim wp2 As WorldPos

        wp2.Map = UserList(UserIndex).flags.TargetMap
        wp2.x = UserList(UserIndex).flags.TargetX
        wp2.Y = UserList(UserIndex).flags.TargetY
    
        If Hechizos(HechizoIndex).NecesitaObj > 0 Then
            If TieneObjEnInv(UserIndex, Hechizos(HechizoIndex).NecesitaObj, Hechizos(HechizoIndex).NecesitaObj2) Then
                PuedeLanzar = True
                'Exit Function
               
            Else
                Call WriteConsoleMsg(UserIndex, "Necesitas un " & ObjData(Hechizos(HechizoIndex).NecesitaObj).name & " para lanzar el hechizo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function

            End If

        End If
    
        If Hechizos(HechizoIndex).CoolDown > 0 Then

            Dim actual            As Long

            Dim segundosFaltantes As Long

            actual = GetTickCount() And &H7FFFFFFF

            If UserList(UserIndex).Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) > actual Then
                segundosFaltantes = Int((UserList(UserIndex).Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) - actual) / 1000)
                Call WriteConsoleMsg(UserIndex, "Debes esperar " & segundosFaltantes & " segundos para volver a tirar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                PuedeLanzar = False
                Exit Function
            Else
                PuedeLanzar = True

            End If

        End If
    
        If UserList(UserIndex).Stats.MinHp > Hechizos(HechizoIndex).RequiredHP Then
            If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
                If UserList(UserIndex).Stats.UserSkills(eSkill.magia) >= Hechizos(HechizoIndex).MinSkill Then
                    If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                        PuedeLanzar = True
                    Else
                        Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                        PuedeLanzar = False

                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo, necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos.", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteLocaleMsg(UserIndex, "221", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
                    PuedeLanzar = False

                End If

            Else
                Call WriteLocaleMsg(UserIndex, "222", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana. Necesitas " & Hechizos(HechizoIndex).ManaRequerido & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No tenes suficiente vida. Necesitas " & Hechizos(HechizoIndex).RequiredHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False

        End If

    Else
        'Call WriteConsoleMsg(UserIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
        Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False

    End If

End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim PosCasteadaX As Integer

    Dim PosCasteadaY As Integer

    Dim PosCasteadaM As Integer

    Dim h            As Integer

    Dim TempX        As Integer

    Dim TempY        As Integer

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
        
    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True

        'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
        For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
            For TempY = PosCasteadaY - 11 To PosCasteadaY + 11

                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.NoDetectable = 0 Then
                            UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 0
                            Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).UserIndex, "Tu invisibilidad ya no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, False))

                        End If

                    End If

                End If

            Next TempY
        Next TempX
    
        Call InfoHechizo(UserIndex)

    End If

End Sub

Sub HechizoSobreArea(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim PosCasteadaX As Byte

    Dim PosCasteadaY As Byte

    Dim PosCasteadaM As Integer

    Dim h            As Integer

    Dim TempX        As Integer

    Dim TempY        As Integer

    PosCasteadaX = UserList(UserIndex).flags.TargetX
    PosCasteadaY = UserList(UserIndex).flags.TargetY
    PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    Dim x         As Long

    Dim Y         As Long
    
    Dim NPCIndex2 As Integer

    Dim Cuantos   As Long
    
    'Envio Palabras magicas, wavs y fxs.
    If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
        Call DecirPalabrasMagicas(h, UserIndex)

    End If
    
    If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
    
        If Hechizos(h).ParticleViaje > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

        End If

    End If
    
    If Hechizos(h).Particle > 0 Then 'Envio Particula?
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

    End If
    
    If Hechizos(h).ParticleViaje = 0 Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))  'Esta linea faltaba. Pablo (ToxicWaste)

    End If

    Dim cuantosuser As Byte

    Dim nameuser    As String
       
    Select Case Hechizos(h).AreaAfecta

        Case 1

            For x = 1 To Hechizos(h).AreaRadio
                For Y = 1 To Hechizos(h).AreaRadio

                    If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex > 0 Then
                        NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex

                        'If NPCIndex2 <> UserIndex Then
                        If UserList(NPCIndex2).flags.Muerto = 0 Then
                                        
                            AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
                            cuantosuser = cuantosuser + 1
                            ' nameuser = nameuser & "," & Npclist(NPCIndex2).Name
                                            
                        End If

                        ' End If
                    End If

                Next
            Next
                    
            ' If cuantosuser > 0 Then
            '     Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & cuantosuser & " usuarios.", FontTypeNames.FONTTYPE_FIGHT)
            ' End If
        Case 2

            For x = 1 To Hechizos(h).AreaRadio
                For Y = 1 To Hechizos(h).AreaRadio

                    If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
                        NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

                        If Npclist(NPCIndex2).Attackable Then
                            AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
                            Cuantos = Cuantos + 1

                        End If

                    End If

                Next
            Next
                
            ' If Cuantos > 0 Then
            '  Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & Cuantos & " criaturas.", FontTypeNames.FONTTYPE_FIGHT)
            '  End If
        Case 3

            For x = 1 To Hechizos(h).AreaRadio
                For Y = 1 To Hechizos(h).AreaRadio

                    If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex > 0 Then
                        NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex

                        'If NPCIndex2 <> UserIndex Then
                        If UserList(NPCIndex2).flags.Muerto = 0 Then
                            AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
                            cuantosuser = cuantosuser + 1

                        End If

                        ' End If
                    End If
                            
                    If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
                        NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

                        If Npclist(NPCIndex2).Attackable Then
                            AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
                            Cuantos = Cuantos + 1
            
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

    b = True

End Sub

Sub HechizoPortal(ByVal UserIndex As Integer, ByRef b As Boolean)

    If UserList(UserIndex).flags.BattleModo = 1 Then
        b = False
        'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
        Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
    Else

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Integer

        Dim uh           As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

        PosCasteadaX = UserList(UserIndex).flags.TargetX
        PosCasteadaY = UserList(UserIndex).flags.TargetY
        PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
        uh = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
        'Envio Palabras magicas, wavs y fxs.
   
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.Map > 0 Or UserList(UserIndex).flags.TargetUser <> 0 Then
            b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)

        Else

            If Hechizos(uh).TeleportX = 1 Then

                If UserList(UserIndex).flags.Portal = 0 Then

                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, False))
            
                    UserList(UserIndex).flags.PortalM = UserList(UserIndex).Pos.Map
                    UserList(UserIndex).flags.PortalX = UserList(UserIndex).flags.TargetX
                    UserList(UserIndex).flags.PortalY = UserList(UserIndex).flags.TargetY
            
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.Intermundia))

                    UserList(UserIndex).accion.AccionPendiente = True
                    UserList(UserIndex).accion.Particula = ParticulasIndex.Runa
                    UserList(UserIndex).accion.TipoAccion = Accion_Barra.Intermundia
                    UserList(UserIndex).accion.HechizoPendiente = uh
            
                    If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
                        Call DecirPalabrasMagicas(uh, UserIndex)

                    End If

                    b = True
                Else
                    Call WriteConsoleMsg(UserIndex, "No podés lanzar mas de un portal a la vez.", FontTypeNames.FONTTYPE_INFO)
                    b = False

                End If

            End If

        End If

    End If

End Sub

Sub HechizoMaterializacion(ByVal UserIndex As Integer, ByRef b As Boolean)

    Dim h   As Integer

    Dim MAT As obj

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
 
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
        b = False
        Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
        ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
    Else
        MAT.Amount = Hechizos(h).MaterializaCant
        MAT.ObjIndex = Hechizos(h).MaterializaObj
        Call MakeObj(MAT, UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
        'Call WriteConsoleMsg(UserIndex, "Has materializado un objeto!!", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
        b = True

    End If

End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 01/10/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
    'usuario
    '***************************************************
    Dim b As Boolean

    Select Case Hechizos(uh).Tipo
        
        Case TipoHechizo.uInvocacion 'Tipo 1

            'Call HechizoInvocacion(UserIndex, b)
        Case TipoHechizo.uEstado 'Tipo 2
            Call HechizoTerrenoEstado(UserIndex, b)

        Case TipoHechizo.uMaterializa 'Tipo 3
            Call HechizoMaterializacion(UserIndex, b)
            
        Case TipoHechizo.uArea 'Tipo 5
            Call HechizoSobreArea(UserIndex, b)
            
        Case TipoHechizo.uPortal 'Tipo 6
            Call HechizoPortal(UserIndex, b)

        Case TipoHechizo.UFamiliar

            ' Call InvocarFamiliar(UserIndex, b)
    End Select

    'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Or UserList(UserIndex).flags.TargetUser <> 0 Then
    '  b = False
    '  Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)

    'Else

    If b Then
        Call SubirSkill(UserIndex, magia)

        'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
        If UserList(UserIndex).clase = eClass.Druid And UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.7
        Else
            UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

        End If

        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
        Call WriteUpdateUserStats(UserIndex)

    End If

End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 01/10/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
    'usuario
    '***************************************************

    Dim b As Boolean

    Select Case Hechizos(uh).Tipo

        Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(UserIndex, b)

        Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
            Call HechizoPropUsuario(UserIndex, b)

        Case TipoHechizo.uCombinados
            Call HechizoCombinados(UserIndex, b)
    
    End Select

    If b Then
        Call SubirSkill(UserIndex, magia)
        'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Hechizos(uh).RequiredHP

        If UserList(UserIndex).Stats.MinHp < 0 Then UserList(UserIndex).Stats.MinHp = 1
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateUserStats(UserList(UserIndex).flags.TargetUser)
        UserList(UserIndex).flags.TargetUser = 0

    End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 01/10/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
    'usuario
    '***************************************************
    Dim b As Boolean

    Select Case Hechizos(uh).Tipo

        Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)

        Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
            Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)

    End Select

    If b Then
        Call SubirSkill(UserIndex, magia)
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

        If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Hechizos(uh).RequiredHP

        If UserList(UserIndex).Stats.MinHp < 0 Then UserList(UserIndex).Stats.MinHp = 1
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

        If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
        Call WriteUpdateUserStats(UserIndex)

    End If

End Sub

Sub LanzarHechizo(Index As Integer, UserIndex As Integer)

    Dim uh As Integer

    uh = UserList(UserIndex).Stats.UserHechizos(Index)

    If PuedeLanzar(UserIndex, uh, Index) Then

        Select Case Hechizos(uh).Target

            Case TargetType.uUsuarios

                If UserList(UserIndex).flags.TargetUser > 0 Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, uh)
                    
                        If Hechizos(uh).CoolDown > 0 Then
                            UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                        End If

                    Else
                        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                End If
        
            Case TargetType.uNPC

                If UserList(UserIndex).flags.TargetNPC > 0 Then
                    If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoNPC(UserIndex, uh)

                        If Hechizos(uh).CoolDown > 0 Then
                            UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF
                    
                        End If
                    
                    Else
                        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                End If
        
            Case TargetType.uUsuariosYnpc

                If UserList(UserIndex).flags.TargetUser > 0 Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(UserIndex, uh)
                    
                        If Hechizos(uh).CoolDown > 0 Then
                            UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                        End If

                    Else
                        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If

                ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then

                    If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
                        If Hechizos(uh).CoolDown > 0 Then
                            UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                        End If

                        Call HandleHechizoNPC(UserIndex, uh)
                    Else
                        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                        ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                End If
        
            Case TargetType.uTerreno

                If Hechizos(uh).CoolDown > 0 Then
                    UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                End If

                Call HandleHechizoTerreno(UserIndex, uh)

        End Select
    
    End If

    If UserList(UserIndex).Counters.Trabajando Then
        Call WriteMacroTrabajoToggle(UserIndex, False)

    End If

    If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
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

    Dim h As Integer, tU As Integer

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tU = UserList(UserIndex).flags.TargetUser

    If Hechizos(h).Invisibilidad = 1 Then
   
        If UserList(tU).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        If UserList(tU).Counters.Saliendo Then
            If UserIndex <> tU Then
                Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
                b = False
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
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)

                End If

            End If

        End If
    
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
            If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                Exit Sub

            End If

        End If
   
        UserList(tU).flags.invisible = 1
        'Ladder
        'Reseteamos el contador de Invisibilidad
        UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
        Call WriteContadores(tU)
        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Envenena > 0 Then
        ' If UserIndex = tU Then
        '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        '   Exit Sub
        'End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Envenenado = Hechizos(h).Envenena
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).desencantar = 1 Then
        ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", FontTypeNames.FONTTYPE_INFOIAO)

        UserList(UserIndex).flags.Envenenado = 0
        UserList(UserIndex).flags.Incinerado = 0
    
        If UserList(UserIndex).flags.Inmovilizado = 1 Then
            UserList(UserIndex).Counters.Inmovilizado = 0
            UserList(UserIndex).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(UserIndex)
            Call FlushBuffer(UserIndex)

        End If
    
        If UserList(UserIndex).flags.Paralizado = 1 Then
            UserList(UserIndex).flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
            Call FlushBuffer(UserIndex)
           
        End If
        
        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)
            Call FlushBuffer(UserIndex)

        End If
    
        If UserList(UserIndex).flags.Maldicion = 1 Then
            UserList(UserIndex).flags.Maldicion = 0
            UserList(UserIndex).Counters.Maldicion = 0

        End If
    
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).incinera = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Incinerado = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).CuraVeneno = 1 Then

        'Verificamos que el usuario no este muerto
        If UserList(tU).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        'Para poder tirar curar veneno a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
        
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
            If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                Exit Sub

            End If

        End If
        
        UserList(tU).flags.Envenenado = 0
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Maldicion = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Maldicion = 1
        UserList(tU).Counters.Maldicion = 200
    
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(tU).flags.Maldicion = 0
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).GolpeCertero = 1 Then
        UserList(tU).flags.GolpeCertero = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Bendicion = 1 Then
        UserList(tU).flags.Bendicion = 1
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Paraliza = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If
            
        Call InfoHechizo(UserIndex)
        b = True

        If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            Call FlushBuffer(tU)
            Exit Sub

        End If
            
        UserList(tU).Counters.Paralisis = Hechizos(h).Duration

        If UserList(tU).flags.Paralizado = 0 Then
            UserList(tU).flags.Paralizado = 1
            Call WriteParalizeOK(tU)
            Call WritePosUpdate(tU)
        End If

    End If

    If Hechizos(h).Velocidad > 0 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If
            
        Call InfoHechizo(UserIndex)
        b = True
                 
        If UserList(tU).Counters.Velocidad = 0 Then
            UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding

        End If

        UserList(tU).Char.speeding = Hechizos(h).Velocidad
        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
        'End If
        UserList(tU).Counters.Velocidad = Hechizos(h).Duration

    End If

    If Hechizos(h).Inmoviliza = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If UserList(tU).flags.Inmovilizado = 1 Then
            Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya esta inmovilizado.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If
            
        Call InfoHechizo(UserIndex)
        b = True
        '  If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
        '   Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
        '   Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
        '   Call FlushBuffer(tU)
        '    Exit Sub
        ' End If
            
        UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

        If UserList(tU).flags.Inmovilizado = 0 Then
            UserList(tU).flags.Inmovilizado = 1
            Call WriteInmovilizaOK(tU)
            Call WritePosUpdate(tU)
            Call FlushBuffer(tU)
        End If

    End If

    If Hechizos(h).RemoverParalisis = 1 Then
        
        'Para poder tirar remo a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)

                End If

            End If

        End If
        
        If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
            Call WriteConsoleMsg(UserIndex, "El objetivo no esta paralizado.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
        
        If UserList(tU).flags.Inmovilizado = 1 Then
            UserList(tU).Counters.Inmovilizado = 0
            UserList(tU).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(tU)
            Call WritePosUpdate(tU)
            ' Call InfoHechizo(UserIndex)
            Call FlushBuffer(tU)

            'b = True
        End If
    
        If UserList(tU).flags.Paralizado = 1 Then
            UserList(tU).flags.Paralizado = 0
            'no need to crypt this
            Call WriteParalizeOK(tU)
            Call FlushBuffer(tU)

            '  b = True
        End If

        b = True
        Call InfoHechizo(UserIndex)

    End If

    If Hechizos(h).RemoverEstupidez = 1 Then
        If UserList(tU).flags.Estupidez = 1 Then

            'Para poder tirar remo estu a un pk en el ring
            If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
                If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                    If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
            UserList(tU).flags.Estupidez = 0
            'no need to crypt this
            Call WriteDumbNoMore(tU)
            Call FlushBuffer(tU)
            Call InfoHechizo(UserIndex)
            b = True

        End If

    End If

    If Hechizos(h).Revivir = 1 Then
        If UserList(tU).flags.Muerto = 1 Then
    
            'No usar resu en mapas con ResuSinEfecto
            'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
            '   Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            '   b = False
            '   Exit Sub
            ' End If
        
            If UserList(tU).accion.TipoAccion = Accion_Barra.Resucitar Then
                Call WriteConsoleMsg(UserIndex, "El usuario ya esta siendo resucitado.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub

            End If
        
            'Para poder tirar revivir a un pk en el ring
            If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
                If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                    If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.Seguro Then
                        'call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub
                    Else
                        Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
                        
            Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, ParticulasIndex.Resucitar, 600, False))
            Call SendData(SendTarget.ToPCArea, tU, PrepareMessageBarFx(UserList(tU).Char.CharIndex, 600, Accion_Barra.Resucitar))
            UserList(tU).accion.AccionPendiente = True
            UserList(tU).accion.Particula = ParticulasIndex.Resucitar
            UserList(tU).accion.TipoAccion = Accion_Barra.Resucitar
                
            'Pablo Toxic Waste (GD: 29/04/07)
            'UserList(tU).Stats.MinAGU = 0
            'UserList(tU).flags.Sed = 1
            'UserList(tU).Stats.MinHam = 0
            'UserList(tU).flags.Hambre = 1
            Call WriteUpdateHungerAndThirst(tU)
            Call InfoHechizo(UserIndex)
            'UserList(tU).Stats.MinMAN = 0
            'UserList(tU).Stats.MinSta = 0
            b = True
            'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        
            'Call RevivirUsuario(tU)
        Else
            b = False

        End If

    End If

    If Hechizos(h).Ceguera = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Ceguera = 1
        UserList(tU).Counters.Ceguera = Hechizos(h).Duration

        Call WriteBlind(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(UserIndex)
        b = True

    End If

    If Hechizos(h).Estupidez = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        If UserList(tU).flags.Estupidez = 0 Then
            UserList(tU).flags.Estupidez = 1
            UserList(tU).Counters.Estupidez = Hechizos(h).Duration

        End If

        Call WriteDumb(tU)
        Call FlushBuffer(tU)

        Call InfoHechizo(UserIndex)
        b = True

    End If

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 04/13/2008
    'Handles the Spells that afect the Stats of an NPC
    '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
    'removidos por users de su misma faccion.
    '***************************************************
    If Hechizos(hIndex).Invisibilidad = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.invisible = 1
        b = True

    End If

    If Hechizos(hIndex).Envenena > 0 Then
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub

        End If

        Call NPCAtacado(NpcIndex, UserIndex)
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
        b = True

    End If

    If Hechizos(hIndex).CuraVeneno = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Envenenado = 0
        b = True

    End If

    If Hechizos(hIndex).RemoverMaldicion = 1 Then
        Call InfoHechizo(UserIndex)
        'Npclist(NpcIndex).flags.Maldicion = 0
        b = True

    End If

    If Hechizos(hIndex).Bendicion = 1 Then
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).flags.Bendicion = 1
        b = True

    End If

    If Hechizos(hIndex).Paraliza = 1 Then
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                b = False
                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, UserIndex)
            Call InfoHechizo(UserIndex)
            Npclist(NpcIndex).flags.Paralizado = 1
            Npclist(NpcIndex).flags.Inmovilizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 2
            b = True
        Else
            'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If

    End If

    If Hechizos(hIndex).RemoverParalisis = 1 Then
        If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                If esArmada(UserIndex) Then
                    Call InfoHechizo(UserIndex)
                    Npclist(NpcIndex).flags.Paralizado = 0
                    Npclist(NpcIndex).Contadores.Paralisis = 0
                    b = True
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If
                
                Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else

                If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    If esCaos(UserIndex) Then
                        Call InfoHechizo(UserIndex)
                        Npclist(NpcIndex).flags.Paralizado = 0
                        Npclist(NpcIndex).Contadores.Paralisis = 0
                        b = True
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        b = False
                        Exit Sub

                    End If

                End If

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If

    End If
 
    If Hechizos(hIndex).Inmoviliza = 1 And Npclist(NpcIndex).flags.Paralizado = 0 Then
        If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                b = False
                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, UserIndex)
            Npclist(NpcIndex).flags.Inmovilizado = 1
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 2
            Call InfoHechizo(UserIndex)
            b = True
        Else
            'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 14/08/2007
    'Handles the Spells that afect the Life NPC
    '14/08/2007 Pablo (ToxicWaste) - Orden general.
    '***************************************************

    Dim daño As Long
    
    'Salud
    If Hechizos(hIndex).SubeHP = 1 Then
        daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
        Call InfoHechizo(UserIndex)
        Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp + daño

        If Npclist(NpcIndex).Stats.MinHp > Npclist(NpcIndex).Stats.MaxHp Then Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MaxHp
        Call WriteConsoleMsg(UserIndex, "Has curado " & daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex, &HFF00))
        b = True
        
    ElseIf Hechizos(hIndex).SubeHP = 2 Then

        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            b = False
            Exit Sub

        End If
        
        Call NPCAtacado(NpcIndex, UserIndex)
        daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
        ' If Hechizos(hIndex).StaffAffected Then
        '     If UserList(UserIndex).clase = eClass.Mage Then
        '         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '             daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
        '             'Aumenta daño segun el staff-
        '             'Daño = (Daño* (70 + BonifBáculo)) / 100
        '         Else
        '             daño = daño * 0.7 'Baja daño a 70% del original
        '         End If
        '     End If
        ' End If
        
        'If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        '    daño = daño * 1.04  'laud magico de los bardos
        'End If
    
        If UserList(UserIndex).flags.DañoMagico > 0 Then
            daño = daño + Porcentaje(daño, UserList(UserIndex).flags.DañoMagico)

        End If
    
        b = True
        
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))

        End If
        
        'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
        daño = daño - Npclist(NpcIndex).Stats.defM
        
        If daño < 0 Then daño = 0
        
        Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
        Call InfoHechizo(UserIndex)
        
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteConsoleMsg(UserIndex, "Le has causado " & daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)

        End If
        
        Call CalcularDarExp(UserIndex, NpcIndex, daño)
    
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex))
    
        If Npclist(NpcIndex).Stats.MinHp < 1 Then
            Npclist(NpcIndex).Stats.MinHp = 0
            Call MuereNpc(NpcIndex, UserIndex)

        End If

    End If

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)

    Dim h As Integer

    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
    If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
        Call DecirPalabrasMagicas(h, UserIndex)

    End If

    If UserList(UserIndex).flags.TargetUser > 0 Then '¿El Hechizo fue tirado sobre un usuario?
        If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
            If Hechizos(h).ParticleViaje > 0 Then
                Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
            Else
                Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

            End If

        End If

        If Hechizos(h).Particle > 0 Then '¿Envio Particula?
            If Hechizos(h).ParticleViaje > 0 Then
                Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
            Else
                Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

            End If

        End If
        
        If Hechizos(h).ParticleViaje = 0 Then
            Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserList(UserIndex).flags.TargetUser).Pos.x, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

        End If
        
        If Hechizos(h).TimeEfect <> 0 Then 'Envio efecto de screen
            Call WriteEfectToScreen(UserIndex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)

        End If

    ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then '¿El Hechizo fue tirado sobre un npc?

        If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
            If Npclist(UserList(UserIndex).flags.TargetNPC).Stats.MinHp < 1 Then

                'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                If Hechizos(h).ParticleViaje > 0 Then
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                Else
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

                End If

            Else

                If Hechizos(h).ParticleViaje > 0 Then
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                Else
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                End If

            End If

        End If
        
        If Hechizos(h).Particle > 0 Then '¿Envio Particula?
            If Npclist(UserList(UserIndex).flags.TargetNPC).Stats.MinHp < 1 Then
                Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.x, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y))
                'Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXToFloor(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.X, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y, Hechizos(H).Particle, Hechizos(H).TimeParticula))
            Else

                If Hechizos(h).ParticleViaje > 0 Then
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                Else
                    Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                End If

            End If

        End If

        If Hechizos(h).ParticleViaje = 0 Then
            Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).wav, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.x, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y))

        End If

    Else ' Entonces debe ser sobre el terreno

        If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
            Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

        End If
        
        If Hechizos(h).Particle > 0 Then 'Envio Particula?
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

        End If
        
        If Hechizos(h).wav <> 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))   'Esta linea faltaba. Pablo (ToxicWaste)

        End If
    
    End If
    
    If UserList(UserIndex).ChatCombate = 1 Then
        If UserList(UserIndex).flags.TargetUser > 0 Then

            'Optimizacion de protocolo por Ladder
            If UserIndex <> UserList(UserIndex).flags.TargetUser Then
                Call WriteConsoleMsg(UserIndex, "HecMSGU*" & h & "*" & UserList(UserList(UserIndex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, "HecMSGA*" & h & "*" & UserList(UserIndex).name, FontTypeNames.FONTTYPE_FIGHT)
    
            Else
                Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            End If

        ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "HecMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

        End If

    End If

End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 02/01/2008
    '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
    '***************************************************

    Dim h As Integer

    Dim daño As Integer

    Dim tempChr           As Integer

    Dim enviarInfoHechizo As Boolean

    enviarInfoHechizo = False
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser
      
    'Hambre
    If Hechizos(h).SubeHam = 1 Then
    
        enviarInfoHechizo = True
    
        daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño

        If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        Call WriteUpdateHungerAndThirst(tempChr)
        b = True
    
    ElseIf Hechizos(h).SubeHam = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
        Else
            Exit Sub

        End If
    
        enviarInfoHechizo = True
    
        daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
        If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        Call WriteUpdateHungerAndThirst(tempChr)
    
        b = True
    
        If UserList(tempChr).Stats.MinHam < 1 Then
            UserList(tempChr).Stats.MinHam = 0
            UserList(tempChr).flags.Hambre = 1

        End If
    
    End If

    'Sed
    If Hechizos(h).SubeSed = 1 Then
    
        enviarInfoHechizo = True
    
        daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño

        If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        b = True
    
    ElseIf Hechizos(h).SubeSed = 2 Then
    
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True
    
        daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1

        End If
    
        b = True

    End If

    ' <-------- Agilidad ---------->
    If Hechizos(h).SubeAgilidad = 1 Then
    
        'Para poder tirar cl a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
    
        enviarInfoHechizo = True
        daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
         
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

        UserList(tempChr).flags.TomoPocion = True
        b = True
        Call WriteFYA(tempChr)
    
    ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True
    
        UserList(tempChr).flags.TomoPocion = True
        daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño < MINATRIBUTOS Then
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        Else
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño

        End If
    
        b = True
        Call WriteFYA(tempChr)

    End If

    ' <-------- Fuerza ---------->
    If Hechizos(h).SubeFuerza = 1 Then

        'Para poder tirar fuerza a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
    
        daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    
        UserList(tempChr).flags.TomoPocion = True
        b = True
    
        enviarInfoHechizo = True
        Call WriteFYA(tempChr)
    ElseIf Hechizos(h).SubeFuerza = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        UserList(tempChr).flags.TomoPocion = True
    
        daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño < MINATRIBUTOS Then
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        Else
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño

        End If

        b = True
        enviarInfoHechizo = True
        Call WriteFYA(tempChr)

    End If

    'Salud
    If Hechizos(h).SubeHP = 1 Then
    
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        'Para poder tirar curar a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
       
        daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
        ' daño = daño + Porcentaje(daño, 2 * UserList(UserIndex).Stats.ELV)
    
        enviarInfoHechizo = True

        UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + daño

        If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex, &HFF00))
    
        b = True
    ElseIf Hechizos(h).SubeHP = 2 Then
    
        If UserIndex = tempChr Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
        '
        ' If Hechizos(H).StaffAffected Then
        '     If UserList(UserIndex).clase = eClass.Mage Then
        '         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '             daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
        '         Else
        '             daño = daño * 0.7 'Baja daño a 70% del original
        '         End If
        '     End If
        ' End If
    
        'If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        '    daño = daño * 1.04  'laud magico de los bardos
        'End If
    
        'cascos antimagia
        'If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        '    daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        'End If
    
        'anillos
        ' If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
        '    daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        'End If

        If UserList(UserIndex).flags.DañoMagico > 0 Then
            daño = daño + Porcentaje(daño, UserList(UserIndex).flags.DañoMagico)

        End If

        If daño < 0 Then daño = 0
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True
    
        'Skill Resistencia Magica
        If UserList(tempChr).Stats.UserSkills(eSkill.Resistencia) > 0 Then

            Dim DefensaMagica As Long

            Dim Absorcion     As Long

            DefensaMagica = UserList(tempChr).Stats.UserSkills(eSkill.Resistencia) / 4
            Absorcion = daño / 100 * DefensaMagica
            daño = daño - Absorcion
        
        End If

        'Defensa Resistencia magica
        If UserList(tempChr).flags.ResistenciaMagica > 0 And Hechizos(h).AntiRm = 0 Then
            daño = daño - Porcentaje(daño, UserList(tempChr).flags.ResistenciaMagica)

        End If
    
        If UserList(tempChr).flags.ResistenciaMagica > 0 And Hechizos(h).AntiRm = 1 Then
            daño = daño + Porcentaje(daño, UserList(tempChr).flags.ResistenciaMagica)

        End If
    
        UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - daño
    
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
        Call SubirSkill(tempChr, Resistencia)
    
        Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex))

        'Muere
        If UserList(tempChr).Stats.MinHp < 1 Then
            'Store it!
            Call Statistics.StoreFrag(UserIndex, tempChr)
            Call ContarMuerte(tempChr, UserIndex)
            UserList(tempChr).Stats.MinHp = 0
            Call ActStats(tempChr, UserIndex)

            '  Call UserDie(tempChr)
        End If
    
        b = True

    End If

    'Mana
    If Hechizos(h).SubeMana = 1 Then
    
        enviarInfoHechizo = True
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño

        If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        b = True
    
    ElseIf Hechizos(h).SubeMana = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño

        If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
        b = True
    
    End If

    'Stamina
    If Hechizos(h).SubeSta = 1 Then
        Call InfoHechizo(UserIndex)
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño

        If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

        End If

        b = True
    ElseIf Hechizos(h).SubeMana = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
        If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
        b = True

    End If

    If enviarInfoHechizo Then
        Call InfoHechizo(UserIndex)

    End If

    Call FlushBuffer(tempChr)

End Sub

Sub HechizoCombinados(ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 02/01/2008
    '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
    '***************************************************

    Dim h As Integer

    Dim daño As Integer

    Dim tempChr           As Integer

    Dim enviarInfoHechizo As Boolean

    enviarInfoHechizo = False
    
    h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    tempChr = UserList(UserIndex).flags.TargetUser
      
    ' <-------- Agilidad ---------->
    If Hechizos(h).SubeAgilidad = 1 Then
    
        'Para poder tirar cl a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
    
        enviarInfoHechizo = True
        daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        'UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
        'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
        
        UserList(tempChr).flags.TomoPocion = True
        b = True
        Call WriteFYA(tempChr)
    
    ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True
    
        UserList(tempChr).flags.TomoPocion = True
        daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño < 6 Then
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        Else
            UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño

        End If

        'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        b = True
        Call WriteFYA(tempChr)

    End If

    ' <-------- Fuerza ---------->
    If Hechizos(h).SubeFuerza = 1 Then

        'Para poder tirar fuerza a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
    
        daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
    
        UserList(tempChr).flags.TomoPocion = True
        b = True
    
        enviarInfoHechizo = True
        Call WriteFYA(tempChr)
    ElseIf Hechizos(h).SubeFuerza = 2 Then

        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        UserList(tempChr).flags.TomoPocion = True
    
        daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
        UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        
        If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño < 6 Then
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        Else
            UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño

        End If
   
        b = True
        enviarInfoHechizo = True
        Call WriteFYA(tempChr)

    End If

    'Salud
    If Hechizos(h).SubeHP = 1 Then
    
        'Verifica que el usuario no este muerto
        If UserList(tempChr).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        'Para poder tirar curar a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
            If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
       
        daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
        'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
        enviarInfoHechizo = True

        UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + daño

        If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
        If UserIndex <> tempChr Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex, &HFF00))
    
        b = True
    ElseIf Hechizos(h).SubeHP = 2 Then
    
        If UserIndex = tempChr Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
        daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        '
        ' If Hechizos(H).StaffAffected Then
        '     If UserList(UserIndex).clase = eClass.Mage Then
        '         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        '             daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
        '         Else
        '             daño = daño * 0.7 'Baja daño a 70% del original
        '         End If
        '     End If
        ' End If
    
        'If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        '    daño = daño * 1.04  'laud magico de los bardos
        'End If
    
        'cascos antimagia
        'If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        '    daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        'End If
    
        'anillos
        ' If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
        '    daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        'End If

        If UserList(UserIndex).flags.DañoMagico > 0 Then
            daño = daño + Porcentaje(daño, UserList(UserIndex).flags.DañoMagico)

        End If

        If daño < 0 Then daño = 0
    
        If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
        If UserIndex <> tempChr Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

        End If
    
        enviarInfoHechizo = True

        'Resistencia Magica By ladder
        If UserList(tempChr).Stats.UserSkills(eSkill.Resistencia) > 0 Then

            Dim DefensaMagica As Long

            Dim Absorcion     As Long

            DefensaMagica = UserList(tempChr).Stats.UserSkills(eSkill.Resistencia) / 4
            Absorcion = daño / 100 * DefensaMagica
            daño = daño - Absorcion

        End If
        
        If UserList(tempChr).flags.ResistenciaMagica > 0 And Hechizos(h).AntiRm = 0 Then
            daño = daño - Porcentaje(daño, UserList(tempChr).flags.ResistenciaMagica)

        End If
        
        'Resistencia Magica By ladder
    
        UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - daño
    
        Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(tempChr, Resistencia)
        Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex))

        'Muere
        If UserList(tempChr).Stats.MinHp < 1 Then
            'Store it!
            Call Statistics.StoreFrag(UserIndex, tempChr)
        
            Call ContarMuerte(tempChr, UserIndex)
            UserList(tempChr).Stats.MinHp = 0
            Call ActStats(tempChr, UserIndex)

            'Call UserDie(tempChr)
        End If
    
        b = True

    End If

    Dim tU As Integer

    tU = tempChr

    If Hechizos(h).Invisibilidad = 1 Then
   
        If UserList(tU).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        If UserList(tU).Counters.Saliendo Then
            If UserIndex <> tU Then
                Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
                b = False
                Exit Sub

            End If

        End If
    
        'Para poder tirar invi a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)

                End If

            End If

        End If
    
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
            If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                Exit Sub

            End If

        End If
   
        UserList(tU).flags.invisible = 1
        'Ladder
        'Reseteamos el contador de Invisibilidad
        UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
        Call WriteContadores(tU)
        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Envenena > 0 Then
        ' If UserIndex = tU Then
        '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        '   Exit Sub
        'End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Envenenado = Hechizos(h).Envenena
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).desencantar = 1 Then
        Call WriteConsoleMsg(UserIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)

        UserList(UserIndex).flags.Envenenado = 0
        UserList(UserIndex).flags.Incinerado = 0
    
        If UserList(UserIndex).flags.Inmovilizado = 1 Then
            UserList(UserIndex).Counters.Inmovilizado = 0
            UserList(UserIndex).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(UserIndex)
            Call FlushBuffer(UserIndex)

        End If
    
        If UserList(UserIndex).flags.Paralizado = 1 Then
            UserList(UserIndex).flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
            Call FlushBuffer(UserIndex)
           
        End If
        
        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)
            Call FlushBuffer(UserIndex)

        End If
    
        If UserList(UserIndex).flags.Maldicion = 1 Then
            UserList(UserIndex).flags.Maldicion = 0
            UserList(UserIndex).Counters.Maldicion = 0

        End If
    
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Sanacion = 1 Then

        UserList(tU).flags.Envenenado = 0
        UserList(tU).flags.Incinerado = 0
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).incinera = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Incinerado = 1
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).CuraVeneno = 1 Then

        'Verificamos que el usuario no este muerto
        If UserList(tU).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub

        End If
    
        'Para poder tirar curar veneno a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else

                    '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                End If

            End If

        End If
        
        'Si sos user, no uses este hechizo con GMS.
        If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
            If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                Exit Sub

            End If

        End If
        
        UserList(tU).flags.Envenenado = 0
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Maldicion = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Maldicion = 1
        UserList(tU).Counters.Maldicion = 200
    
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(tU).flags.Maldicion = 0
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).GolpeCertero = 1 Then
        UserList(tU).flags.GolpeCertero = 1
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Bendicion = 1 Then
        UserList(tU).flags.Bendicion = 1
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Paraliza = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If
            
        enviarInfoHechizo = True
        b = True

        If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            Call FlushBuffer(tU)
            Exit Sub

        End If
            
        UserList(tU).Counters.Paralisis = Hechizos(h).Duration

        If UserList(tU).flags.Paralizado = 0 Then
            UserList(tU).flags.Paralizado = 1
            Call WriteParalizeOK(tU)
            Call WritePosUpdate(tU)
        End If

    End If

    If Hechizos(h).Inmoviliza = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If
            
        enviarInfoHechizo = True
        b = True
            
        UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

        If UserList(tU).flags.Inmovilizado = 0 Then
            UserList(tU).flags.Inmovilizado = 1
            Call WriteInmovilizaOK(tU)
            Call WritePosUpdate(tU)
            Call FlushBuffer(tU)

        End If

    End If

    If Hechizos(h).RemoverParalisis = 1 Then
        
        'Para poder tirar remo a un pk en el ring
        If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
            If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
                If esArmada(UserIndex) Then
                    'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub

                End If

                If UserList(UserIndex).flags.Seguro Then
                    'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                Else
                    Call VolverCriminal(UserIndex)

                End If

            End If
            
        End If

        If UserList(tU).flags.Inmovilizado = 1 Then
            UserList(tU).Counters.Inmovilizado = 0
            UserList(tU).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(tU)
            enviarInfoHechizo = True
            Call FlushBuffer(tU)
            b = True

        End If

        If UserList(tU).flags.Paralizado = 1 Then
            UserList(tU).flags.Paralizado = 0
            'no need to crypt this
            Call WriteParalizeOK(tU)
            enviarInfoHechizo = True
            Call FlushBuffer(tU)
            b = True

        End If

    End If

    If Hechizos(h).Ceguera = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        UserList(tU).flags.Ceguera = 1
        UserList(tU).Counters.Ceguera = Hechizos(h).Duration

        Call WriteBlind(tU)
        Call FlushBuffer(tU)
        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Estupidez = 1 Then
        If UserIndex = tU Then
            'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If

        If UserList(tU).flags.Estupidez = 0 Then
            UserList(tU).flags.Estupidez = 1
            UserList(tU).Counters.Estupidez = Hechizos(h).Duration

        End If

        Call WriteDumb(tU)
        Call FlushBuffer(tU)

        enviarInfoHechizo = True
        b = True

    End If

    If Hechizos(h).Velocidad > 0 Then

        If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
        If UserIndex <> tU Then
            Call UsuarioAtacadoPorUsuario(UserIndex, tU)

        End If
            
        enviarInfoHechizo = True
        b = True
            
        If UserList(tU).Counters.Velocidad = 0 Then
            UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding

        End If

        UserList(tU).Char.speeding = Hechizos(h).Velocidad
        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            
        UserList(tU).Counters.Velocidad = Hechizos(h).Duration

    End If

    If enviarInfoHechizo Then
        Call InfoHechizo(UserIndex)

    End If

    Call FlushBuffer(tempChr)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)

    'Call LogTarea("Sub UpdateUserHechizos")

    Dim LoopC As Byte

    'Actualiza un solo slot
    If Not UpdateAll Then

        'Actualiza el inventario
        If UserList(UserIndex).Stats.UserHechizos(slot) > 0 Then
            Call ChangeUserHechizo(UserIndex, slot, UserList(UserIndex).Stats.UserHechizos(slot))
        Else
            Call ChangeUserHechizo(UserIndex, slot, 0)

        End If

    Else

        'Actualiza todos los slots
        For LoopC = 1 To MAXUSERHECHIZOS

            'Actualiza el inventario
            If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
                Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
            Else
                Call ChangeUserHechizo(UserIndex, LoopC, 0)

            End If

        Next LoopC

    End If

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Hechizo As Integer)

    'Call LogTarea("ChangeUserHechizo")
    
    UserList(UserIndex).Stats.UserHechizos(slot) = Hechizo
    
    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call WriteChangeSpellSlot(UserIndex, slot)
    Else
        Call WriteChangeSpellSlot(UserIndex, slot)

    End If

End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

    If (Dire <> 1 And Dire <> -1) Then Exit Sub
    If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    If Dire = 1 Then 'Mover arriba
        If CualHechizo = 1 Then
            Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

            'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
            If UserList(UserIndex).flags.Hechizo > 0 Then
                UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1

            End If

        End If

    Else 'mover abajo

        If CualHechizo = MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
            UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

            'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
            If UserList(UserIndex).flags.Hechizo > 0 Then
                UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1

            End If

        End If

    End If

End Sub

Sub AreaHechizo(UserIndex As Integer, NpcIndex As Integer, x As Byte, Y As Byte, npc As Boolean)

    Dim calculo      As Integer

    Dim TilesDifUser As Integer

    Dim TilesDifNpc  As Integer

    Dim tilDif       As Integer

    Dim h2           As Integer

    Dim Hit          As Integer

    Dim daño As Integer

    Dim porcentajeDesc As Integer

    h2 = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

    'Calculo de descuesto de golpe por cercania.
    TilesDifUser = x + Y

    If npc Then
        If Hechizos(h2).SubeHP = 2 Then
            TilesDifNpc = Npclist(NpcIndex).Pos.x + Npclist(NpcIndex).Pos.Y
            
            tilDif = TilesDifUser - TilesDifNpc
            
            Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

            If tilDif <> 0 Then
                porcentajeDesc = Abs(tilDif) * 20
                daño = Hit / 100 * porcentajeDesc
                daño = Hit - daño
            Else
                daño = Hit

            End If
            
            If UserList(UserIndex).flags.DañoMagico > 0 Then
                daño = daño + Porcentaje(daño, UserList(UserIndex).flags.DañoMagico)

            End If
            
            Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
            
            If UserList(UserIndex).ChatCombate = 1 Then
                Call WriteConsoleMsg(UserIndex, "Le has causado " & daño & " puntos de daño a " & Npclist(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call CalcularDarExp(UserIndex, NpcIndex, daño)
                
            If Npclist(NpcIndex).Stats.MinHp <= 0 Then
                'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Npclist(NpcIndex).GiveEXP
                'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Npclist(NpcIndex).GiveGLD
                Call MuereNpc(NpcIndex, UserIndex)

            End If

            Exit Sub

        End If

    Else

        TilesDifNpc = UserList(NpcIndex).Pos.x + UserList(NpcIndex).Pos.Y
        tilDif = TilesDifUser - TilesDifNpc

        If Hechizos(h2).SubeHP = 2 Then
            If UserIndex = NpcIndex Then
                Exit Sub

            End If

            If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
            If UserIndex <> NpcIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                
            Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

            If tilDif <> 0 Then
                porcentajeDesc = Abs(tilDif) * 20
                daño = Hit / 100 * porcentajeDesc
                daño = Hit - daño
            Else
                daño = Hit

            End If
                        
            If UserList(UserIndex).flags.DañoMagico > 0 Then
                daño = daño + Porcentaje(daño, UserList(UserIndex).flags.DañoMagico)

            End If
                            
            If Hechizos(h2).AntiRm = 1 Then
                            
                If UserList(NpcIndex).Stats.UserSkills(eSkill.Resistencia) > 0 Then

                    Dim DefensaMagica As Long

                    Dim Absorcion     As Long

                    DefensaMagica = UserList(NpcIndex).Stats.UserSkills(eSkill.Resistencia) / 4
                    Absorcion = daño / 100 * DefensaMagica
                    daño = daño - Absorcion

                End If

            End If
                        
            UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp - daño
                    
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call SubirSkill(NpcIndex, Resistencia)
            Call WriteUpdateUserStats(NpcIndex)
                
            'Muere
            If UserList(NpcIndex).Stats.MinHp < 1 Then
                'Store it!
                Call Statistics.StoreFrag(UserIndex, NpcIndex)
                        
                Call ContarMuerte(NpcIndex, UserIndex)
                UserList(NpcIndex).Stats.MinHp = 0
                Call ActStats(NpcIndex, UserIndex)

                'Call UserDie(NpcIndex)
            End If

        End If
                
        If Hechizos(h2).SubeHP = 1 Then
            If (TriggerZonaPelea(UserIndex, NpcIndex) <> TRIGGER6_PERMITE) Then
                If Status(UserIndex) = 1 And Status(NpcIndex) <> 1 Then
                    Exit Sub

                End If

            End If

            Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

            If tilDif <> 0 Then
                porcentajeDesc = Abs(tilDif) * 20
                daño = Hit / 100 * porcentajeDesc
                daño = Hit - daño
            Else
                daño = Hit

            End If
 
            UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp + daño

            If UserList(NpcIndex).Stats.MinHp > UserList(NpcIndex).Stats.MaxHp Then UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MaxHp

        End If
 
        If UserIndex <> NpcIndex Then
            Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

        End If
                    
        Call WriteUpdateUserStats(NpcIndex)

    End If
                
    If Hechizos(h2).Envenena > 0 Then
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
                    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If
                    
        UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
        Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha envenenado.", FontTypeNames.FONTTYPE_FIGHT)

    End If
                
    If Hechizos(h2).Paraliza = 1 Then
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If
            
        Call WriteConsoleMsg(NpcIndex, "Has sido paralizado.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

        If UserList(NpcIndex).flags.Paralizado = 0 Then
            UserList(NpcIndex).flags.Paralizado = 1
            Call WriteParalizeOK(NpcIndex)
            Call FlushBuffer(NpcIndex)

        End If
            
    End If
                
    If Hechizos(h2).Inmoviliza = 1 Then
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If
                    
        Call WriteConsoleMsg(NpcIndex, "Has sido inmovilizado.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration

        If UserList(NpcIndex).flags.Inmovilizado = 0 Then
            UserList(NpcIndex).flags.Inmovilizado = 1
            Call WriteInmovilizaOK(NpcIndex)
            Call WritePosUpdate(NpcIndex)
            Call FlushBuffer(NpcIndex)
        End If

    End If
                
    If Hechizos(h2).Ceguera = 1 Then
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If
                    
        UserList(NpcIndex).flags.Ceguera = 1
        UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
        Call WriteConsoleMsg(NpcIndex, "Te han cegado.", FontTypeNames.FONTTYPE_INFO)
            
        Call WriteBlind(NpcIndex)
        Call FlushBuffer(NpcIndex)

    End If
                
    If Hechizos(h2).Velocidad > 0 Then
    
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If

        If UserList(NpcIndex).Counters.Velocidad = 0 Then
            UserList(NpcIndex).flags.VelocidadBackup = UserList(NpcIndex).Char.speeding

        End If

        UserList(NpcIndex).Char.speeding = Hechizos(h2).Velocidad
        Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSpeedingACT(UserList(NpcIndex).Char.CharIndex, UserList(NpcIndex).Char.speeding))
        UserList(NpcIndex).Counters.Velocidad = Hechizos(h2).Duration

    End If
                
    If Hechizos(h2).Maldicion = 1 Then
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If

        Call WriteConsoleMsg(NpcIndex, "Ahora estas maldito. No podras Atacar", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Maldicion = 1
        UserList(NpcIndex).Counters.Maldicion = Hechizos(h2).Duration

    End If
                
    If Hechizos(h2).RemoverMaldicion = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Te han removido la maldicion.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Maldicion = 0

    End If
                
    If Hechizos(h2).GolpeCertero = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Tu proximo golpe sera certero.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.GolpeCertero = 1

    End If
                
    If Hechizos(h2).Bendicion = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Has sido bendecido.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Bendicion = 1

    End If
                  
    If Hechizos(h2).incinera = 1 Then
        If UserIndex = NpcIndex Then
            Exit Sub

        End If
    
        If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
        If UserIndex <> NpcIndex Then
            Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

        End If

        UserList(NpcIndex).flags.Incinerado = 1
        Call WriteConsoleMsg(NpcIndex, "Has sido Incinerado.", FontTypeNames.FONTTYPE_INFO)

    End If
                
    If Hechizos(h2).Invisibilidad = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Ahora sos invisible.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.invisible = 1
        UserList(NpcIndex).Counters.Invisibilidad = Hechizos(h2).Duration
        Call WriteContadores(NpcIndex)
        Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSetInvisible(UserList(NpcIndex).Char.CharIndex, True))

    End If
                              
    If Hechizos(h2).Sanacion = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Has sido sanado.", FontTypeNames.FONTTYPE_INFO)
        UserList(NpcIndex).flags.Envenenado = 0
        UserList(NpcIndex).flags.Incinerado = 0

    End If
                
    If Hechizos(h2).RemoverParalisis = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Has sido removido.", FontTypeNames.FONTTYPE_INFO)

        If UserList(NpcIndex).flags.Inmovilizado = 1 Then
            UserList(NpcIndex).Counters.Inmovilizado = 0
            UserList(NpcIndex).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(NpcIndex)
            Call FlushBuffer(NpcIndex)

        End If

        If UserList(NpcIndex).flags.Paralizado = 1 Then
            UserList(NpcIndex).flags.Paralizado = 0
            'no need to crypt this
            Call WriteParalizeOK(NpcIndex)
            Call FlushBuffer(NpcIndex)

        End If

    End If
                
    If Hechizos(h2).desencantar = 1 Then
        Call WriteConsoleMsg(NpcIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)
                    
        UserList(NpcIndex).flags.Envenenado = 0
        UserList(NpcIndex).flags.Incinerado = 0
                    
        If UserList(NpcIndex).flags.Inmovilizado = 1 Then
            UserList(NpcIndex).Counters.Inmovilizado = 0
            UserList(NpcIndex).flags.Inmovilizado = 0
            Call WriteInmovilizaOK(NpcIndex)
            Call FlushBuffer(NpcIndex)

        End If
                    
        If UserList(NpcIndex).flags.Paralizado = 1 Then
            UserList(NpcIndex).flags.Paralizado = 0
            Call WriteParalizeOK(NpcIndex)
            Call FlushBuffer(NpcIndex)
                       
        End If
                    
        If UserList(NpcIndex).flags.Ceguera = 1 Then
            UserList(NpcIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(NpcIndex)
            Call FlushBuffer(NpcIndex)

        End If
                    
        If UserList(NpcIndex).flags.Maldicion = 1 Then
            UserList(NpcIndex).flags.Maldicion = 0
            UserList(NpcIndex).Counters.Maldicion = 0

        End If

    End If
        
End Sub
