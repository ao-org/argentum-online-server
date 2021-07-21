Attribute VB_Name = "Trabajo"
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


Public Sub Trabajar(ByVal UserIndex As Integer, ByVal Skill As eSkill)
    Dim DummyInt As Integer
    With UserList(UserIndex)
    Select Case Skill
        Case eSkill.Pescar
288                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
290                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                    'Check interval
292                 'If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

294                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 1      ' Subtipo: Caña de Pescar

296                         If (MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).Blocked And FLAG_AGUA) <> 0 Then
298                             If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X + 1, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X, .Pos.Y + 1).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X - 1, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).Blocked And FLAG_AGUA) <> 0 Then

300                                 Call DoPescar(UserIndex, False)
                                   
                                Else
304                                 Call WriteConsoleMsg(UserIndex, "Acércate a la costa para pescar.", FontTypeNames.FONTTYPE_INFO)
306                                 Call WriteMacroTrabajoToggle(UserIndex, False)

                                End If
                            
                            Else
308                             Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
310                             Call WriteMacroTrabajoToggle(UserIndex, False)
    
                            End If
                    
312                     Case 2      ' Subtipo: Red de Pesca
    
314                         If (MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).Blocked And FLAG_AGUA) <> 0 Then
                            
316                             If Abs(.Pos.X - .Trabajo.Target_X) + Abs(.Pos.Y - .Trabajo.Target_Y) > 8 Then
318                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
320                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                
322                             If UserList(UserIndex).Stats.UserSkills(eSkill.Pescar) < 80 Then
324                                 Call WriteConsoleMsg(UserIndex, "Para utilizar la red de pesca debes tener 80 skills en recoleccion.", FontTypeNames.FONTTYPE_INFO)
326                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                    
328                             If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
330                                 Call WriteConsoleMsg(UserIndex, "Esta prohibida la pesca masiva en las ciudades.", FontTypeNames.FONTTYPE_INFO)
332                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                    
334                             If UserList(UserIndex).flags.Navegando = 0 Then
336                                 Call WriteConsoleMsg(UserIndex, "Necesitas estar sobre tu barca para utilizar la red de pesca.", FontTypeNames.FONTTYPE_INFO)
338                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                    
340                             Call DoPescar(UserIndex, True)
342                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                        
                            Else
                        
344                             Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
346                             Call WriteWorkRequestTarget(UserIndex, 0)
    
                            End If
                        End Select
        Case eSkill.Mineria
            
454                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
456                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                    'Check interval
458                 If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

460                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 8  ' Herramientas de Mineria - Piquete
                
                            'Target whatever is in the tile
462                         Call LookatTile(UserIndex, .Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y)
                            
464                         DummyInt = MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex
                            
466                         If DummyInt > 0 Then

                                'Check distance
468                             If Abs(.Pos.X - .Trabajo.Target_X) + Abs(.Pos.Y - .Trabajo.Target_Y) > 2 Then
470                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
472                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If

                                '¡Hay un yacimiento donde clickeo?
474                             If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then

                                    ' Si el Yacimiento requiere herramienta `Dorada` y la herramienta no lo es, o vice versa.
                                    ' Se usa para el yacimiento de Oro.
476                                 If ObjData(DummyInt).Dorada <> ObjData(.Invent.HerramientaEqpObjIndex).Dorada Then
478                                     Call WriteConsoleMsg(UserIndex, "El pico dorado solo puede extraer minerales del yacimiento de Oro.", FontTypeNames.FONTTYPE_INFO)
480                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If

482                                 If MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount <= 0 Then
484                                     Call WriteConsoleMsg(UserIndex, "Este yacimiento no tiene mas minerales para entregar.", FontTypeNames.FONTTYPE_INFO)
486                                     Call WriteWorkRequestTarget(UserIndex, 0)
488                                     Call WriteMacroTrabajoToggle(UserIndex, False)
                                        Exit Sub

                                    End If

490                                 Call DoMineria(UserIndex, .Trabajo.Target_X, .Trabajo.Target_Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                                Else
492                                 Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
494                                 Call WriteWorkRequestTarget(UserIndex, 0)

                                End If

                            Else
496                             Call WriteConsoleMsg(UserIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
498                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                    End Select
        Case eSkill.Talar
350                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub

352                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
        
                    'Check interval
354                 If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

356                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 6      ' Herramientas de Carpinteria - Hacha

                            ' Ahora se puede talar en la ciudad
                            'If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                            '    Call WriteConsoleMsg(UserIndex, "Esta prohibido talar arboles en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                            '    Call WriteWorkRequestTarget(UserIndex, 0)
                            '    Exit Sub
                            'End If
                            
358                         DummyInt = MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex
                            
360                         If DummyInt > 0 Then
362                             If Abs(.Pos.X - .Trabajo.Target_X) + Abs(.Pos.Y - .Trabajo.Target_Y) > 1 Then
364                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
366                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
368                             If .Pos.X = .Trabajo.Target_X And .Pos.Y = .Trabajo.Target_Y Then
370                                 Call WriteConsoleMsg(UserIndex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
372                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If

374                             If ObjData(DummyInt).Elfico <> ObjData(.Invent.HerramientaEqpObjIndex).Elfico Then
376                                 Call WriteConsoleMsg(UserIndex, "Sólo puedes talar árboles elficos con un hacha élfica.", FontTypeNames.FONTTYPE_INFO)
378                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If

380                             If MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount <= 0 Then
382                                 Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas leña.", FontTypeNames.FONTTYPE_INFO)
384                                 Call WriteWorkRequestTarget(UserIndex, 0)
386                                 Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Exit Sub

                                End If

                                '¡Hay un arbol donde clickeo?
388                             If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
390                                 Call DoTalar(UserIndex, .Trabajo.Target_X, .Trabajo.Target_Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                                End If

                            Else
392                             Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
394                             Call WriteWorkRequestTarget(UserIndex, 0)

396                             If UserList(UserIndex).Counters.Trabajando > 1 Then
398                                 Call WriteMacroTrabajoToggle(UserIndex, False)

                                End If

                            End If
                
                    End Select
    End Select
    End With
End Sub

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
        '********************************************************
        'Autor: Nacho (Integer)
        'Last Modif: 28/01/2007
        'Chequea si ya debe mostrarse
        'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
        '********************************************************
        
        On Error GoTo DoPermanecerOculto_Err
    
100     With UserList(UserIndex)

            ' WyroX: Si tiene armadura de cazador, no se le va nunca lo oculto
102         If .clase = eClass.Hunter And TieneArmaduraCazador(UserIndex) Then Exit Sub
    
104         .Counters.TiempoOculto = .Counters.TiempoOculto - 1

106         If .Counters.TiempoOculto <= 0 Then

108             .Counters.TiempoOculto = 0
110             .flags.Oculto = 0

112             If .flags.Navegando = 1 Then
            
114                 If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
116                     Call EquiparBarco(UserIndex)
124                     Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
126                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
                        Call RefreshCharStatus(UserIndex)
                    End If

                Else

128                 If .flags.invisible = 0 And .flags.AdminInvisible = 0 Then
130                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
132                     Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
            End If
    
        End With

        Exit Sub

DoPermanecerOculto_Err:
134     Call TraceError(Err.Number, Err.Description, "Trabajo.DoPermanecerOculto", Erl)

136
        
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

        'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
        'Modifique la fórmula y ahora anda bien.
        On Error GoTo ErrHandler

        Dim Suerte As Double
        Dim res    As Integer
        Dim Skill  As Integer
    
100     With UserList(UserIndex)

102         If .flags.Navegando = 1 And .clase <> eClass.Pirat Then
104             Call WriteLocaleMsg(UserIndex, "56", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
    
106         Skill = .Stats.UserSkills(eSkill.Ocultarse)
108         Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
110         res = RandomNumber(1, 100)

112         If res <= Suerte Then

114             .flags.Oculto = 1
116             Suerte = (-0.000001 * (100 - Skill) ^ 3)
118             Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
120             Suerte = Suerte + (-0.0088 * (100 - Skill))
122             Suerte = Suerte + (0.9571)
124             Suerte = Suerte * IntervaloOculto
        
126             If .clase = eClass.Bandit Then
128                 .Counters.TiempoOculto = Int(Suerte / 2)
                Else
130                 .Counters.TiempoOculto = Suerte
                End If
    
132             If .flags.AnilloOcultismo = 1 Then
134                 .Counters.TiempoOculto = Suerte * 3
                Else
136                 .Counters.TiempoOculto = Suerte
                End If
  
138             If .flags.Navegando = 1 Then
140                 If .clase = eClass.Pirat Then
142                     .Char.Body = iFragataFantasmal
144                     .flags.Oculto = 1
146                     .Counters.TiempoOculto = IntervaloOculto
                         
148                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
150                     Call WriteConsoleMsg(UserIndex, "¡Te has camuflado como barco fantasma!", FontTypeNames.FONTTYPE_INFO)
                        Call RefreshCharStatus(UserIndex)
                    End If
                Else
152                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
                    
                    'Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
154                 Call WriteLocaleMsg(UserIndex, "55", FontTypeNames.FONTTYPE_INFO)
                End If


156             Call SubirSkill(UserIndex, Ocultarse)
            Else

158             If Not .flags.UltimoMensaje = 4 Then
                    'Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
160                 Call WriteLocaleMsg(UserIndex, "57", FontTypeNames.FONTTYPE_INFO)
162                 .flags.UltimoMensaje = 4
                End If

            End If

164         .Counters.Ocultando = .Counters.Ocultando + 1
    
        End With

        Exit Sub

ErrHandler:
166     Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer)
        
        On Error GoTo DoNavega_Err

100     With UserList(UserIndex)

102         If .Invent.BarcoObjIndex <> .Invent.Object(Slot).ObjIndex Then

104             If Not EsGM(UserIndex) Then
            
106                 Select Case Barco.Subtipo
        
                        Case 2  'Galera
        
108                         If .clase <> eClass.Trabajador And .clase <> eClass.Pirat Then
110                             Call WriteConsoleMsg(UserIndex, "¡Solo Piratas y trabajadores pueden usar galera!", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
112                     Case 3  'Galeón
                    
114                         If .clase <> eClass.Pirat Then
116                             Call WriteConsoleMsg(UserIndex, "Solo los Piratas pueden usar Galeón!!", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
                    End Select
                    
                End If
            
                Dim SkillNecesario As Byte
                
                If .Invent.Object(Slot).ObjIndex = 200 Or .Invent.Object(Slot).ObjIndex = 199 Then 'Traje nw alto y bajo
                    SkillNecesario = 0
                Else
118                 SkillNecesario = IIf(.clase = eClass.Trabajador Or .clase = eClass.Pirat, Barco.MinSkill \ 2, Barco.MinSkill)
                End If
                ' Tiene el skill necesario?
120             If .Stats.UserSkills(eSkill.Navegacion) < SkillNecesario Then
122                 Call WriteConsoleMsg(UserIndex, "Necesitas al menos " & SkillNecesario & " puntos en navegación para poder usar este " & IIf(Barco.Subtipo = 0, "traje", "barco") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
124             If .Invent.BarcoObjIndex = 0 Then
126                 Call WriteNavigateToggle(UserIndex)
128                 .flags.Navegando = 1
                End If
    
130             .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
132             .Invent.BarcoSlot = Slot
    
134             If .flags.Montado > 0 Then
136                 Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
                End If
                
138             If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
140                 Call WriteConsoleMsg(UserIndex, "Pierdes el efecto del mimetismo.", FontTypeNames.FONTTYPE_INFO)
142                 .Counters.Mimetismo = 0
144                 .flags.Mimetizado = e_EstadoMimetismo.Desactivado
                    Call RefreshCharStatus(UserIndex)
                End If
                
                If .flags.invisible = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Pierdes el efecto de la invisibilidad.", FontTypeNames.FONTTYPE_INFO)
                    .flags.invisible = 0
                    .Counters.Invisibilidad = 0
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
                End If
    
146             Call EquiparBarco(UserIndex)
            
            Else
148             Call WriteNadarToggle(UserIndex, False)
            
150             Call WriteNavigateToggle(UserIndex)
    
152             .flags.Navegando = 0
154             .Invent.BarcoObjIndex = 0
156             .Invent.BarcoSlot = 0
    
158             If .flags.Muerto = 0 Then
160                 .Char.Head = .OrigChar.Head
        
162                 If .Invent.ArmourEqpObjIndex > 0 Then
164                     .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                    Else
166                     Call DarCuerpoDesnudo(UserIndex)
                    End If
        
168                 If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

170                 If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

172                 If .Invent.NudilloObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.NudilloObjIndex).WeaponAnim

174                 If .Invent.HerramientaEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.HerramientaEqpObjIndex).WeaponAnim

176                 If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim

                Else
178                 .Char.Body = iCuerpoMuerto
180                 .Char.Head = 0
182                 .Char.ShieldAnim = NingunEscudo
184                 .Char.WeaponAnim = NingunArma
186                 .Char.CascoAnim = NingunCasco

                End If

            End If
            
188         Call ActualizarVelocidadDeUsuario(UserIndex)
        
            ' Volver visible
190         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 And .flags.invisible = 0 Then
192             .flags.Oculto = 0
194             .Counters.TiempoOculto = 0

                'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
196             Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
198             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            End If

200         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
202         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, .Pos.X, .Pos.Y))
    
        End With
        
        Exit Sub

DoNavega_Err:
204     Call TraceError(Err.Number, Err.Description, "Trabajo.DoNavega", Erl)

206
        
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
        
        On Error GoTo FundirMineral_Err

100     If UserList(UserIndex).clase <> eClass.Trabajador Then
102         Call WriteConsoleMsg(UserIndex, "Tu clase no tiene el conocimiento suficiente para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
104     If UserList(UserIndex).flags.Privilegios And (PlayerType.Consejero) Then
            Exit Sub
        End If

106     If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then

            Dim SkillRequerido As Integer
108         SkillRequerido = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill
   
110         If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
                UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= SkillRequerido Then
            
112             Call DoLingotes(UserIndex)
        
114         ElseIf SkillRequerido > 100 Then
116             Call WriteConsoleMsg(UserIndex, "Los mortales no pueden fundir este mineral.", FontTypeNames.FONTTYPE_INFO)
                
            Else
118             Call WriteConsoleMsg(UserIndex, "No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas " & SkillRequerido & " puntos en minería.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

FundirMineral_Err:
120     Call TraceError(Err.Number, Err.Description, "Trabajo.FundirMineral", Erl)
122
        
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        'Call LogTarea("Sub TieneObjetos")
        
        On Error GoTo TieneObjetos_Err
        

        Dim i     As Long

        Dim Total As Long

100     For i = 1 To UserList(UserIndex).CurrentInventorySlots

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
104             Total = Total + UserList(UserIndex).Invent.Object(i).amount

            End If

106     Next i

108     If cant <= Total Then
110         TieneObjetos = True
            Exit Function

        End If
        
        
        Exit Function

TieneObjetos_Err:
112     Call TraceError(Err.Number, Err.Description, "Trabajo.TieneObjetos", Erl)
114
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        'Call LogTarea("Sub QuitarObjetos")
        
        On Error GoTo QuitarObjetos_Err
        
100     With UserList(UserIndex)

            Dim i As Long
    
102         For i = 1 To .CurrentInventorySlots
    
104             If .Invent.Object(i).ObjIndex = ItemIndex Then
    
106                 .Invent.Object(i).amount = .Invent.Object(i).amount - cant
    
108                 If .Invent.Object(i).amount <= 0 Then
110                     If .Invent.Object(i).Equipped Then
112                         Call Desequipar(UserIndex, i)
                        End If
    
114                     cant = Abs(.Invent.Object(i).amount)
116                     .Invent.Object(i).amount = 0
118                     .Invent.Object(i).ObjIndex = 0
                    Else
120                     cant = 0
    
                    End If

122                 Call UpdateUserInv(False, UserIndex, i)
            
124                 If cant = 0 Then
126                     QuitarObjetos = True
                        Exit Function
                    End If
    
                End If
    
128         Next i

        End With
        
        Exit Function

QuitarObjetos_Err:
130     Call TraceError(Err.Number, Err.Description, "Trabajo.QuitarObjetos", Erl)
132
        
End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
102     If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
104     If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)

        
        Exit Sub

HerreroQuitarMateriales_Err:
106     Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroQuitarMateriales", Erl)
108
        
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)

102     If ObjData(ItemIndex).MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, ObjData(ItemIndex).MaderaElfica, UserIndex)

        
        Exit Sub

CarpinteroQuitarMateriales_Err:
104     Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroQuitarMateriales", Erl)
106
        
End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)

        
        Exit Sub

AlquimistaQuitarMateriales_Err:
102     Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaQuitarMateriales", Erl)
104
        
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreQuitarMateriales_Err
        

100     If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex)
102     If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
104     If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex)

        
        Exit Sub

SastreQuitarMateriales_Err:
106     Call TraceError(Err.Number, Err.Description, "Trabajo.SastreQuitarMateriales", Erl)
108
        
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo CarpinteroTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).Madera > 0 Then
102         If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
106             CarpinteroTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
        
110     If ObjData(ItemIndex).MaderaElfica > 0 Then
112         If Not TieneObjetos(LeñaElfica, ObjData(ItemIndex).MaderaElfica, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "No tenes suficiente madera elfica.", FontTypeNames.FONTTYPE_INFO)
116             CarpinteroTieneMateriales = False
118             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If

        End If
    
120     CarpinteroTieneMateriales = True

        
        Exit Function

CarpinteroTieneMateriales_Err:
122     Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroTieneMateriales", Erl)
124
        
End Function

Function AlquimistaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo AlquimistaTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).Raices > 0 Then
102         If Not TieneObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes raices.", FontTypeNames.FONTTYPE_INFO)
106             AlquimistaTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
    
110     AlquimistaTieneMateriales = True

        
        Exit Function

AlquimistaTieneMateriales_Err:
112     Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaTieneMateriales", Erl)
114
        
End Function

Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo SastreTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).PielLobo > 0 Then
102         If Not TieneObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de lobo.", FontTypeNames.FONTTYPE_INFO)
106             SastreTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
    
110     If ObjData(ItemIndex).PielOsoPardo > 0 Then
112         If Not TieneObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso pardo.", FontTypeNames.FONTTYPE_INFO)
116             SastreTieneMateriales = False
118             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
    
120     If ObjData(ItemIndex).PielOsoPolaR > 0 Then
122         If Not TieneObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex) Then
124             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso polar.", FontTypeNames.FONTTYPE_INFO)
126             SastreTieneMateriales = False
128             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
    
130     SastreTieneMateriales = True

        
        Exit Function

SastreTieneMateriales_Err:
132     Call TraceError(Err.Number, Err.Description, "Trabajo.SastreTieneMateriales", Erl)
134
        
End Function

Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo HerreroTieneMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then
102         If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
106             HerreroTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If

110     If ObjData(ItemIndex).LingP > 0 Then
112         If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
116             HerreroTieneMateriales = False
118             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If

120     If ObjData(ItemIndex).LingO > 0 Then
122         If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
124             Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
126             HerreroTieneMateriales = False
128             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If

130     HerreroTieneMateriales = True

        
        Exit Function

HerreroTieneMateriales_Err:
132     Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroTieneMateriales", Erl)
134
        
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruir_Err
        
100     PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkHerreria

        
        Exit Function

PuedeConstruir_Err:
102     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruir", Erl)
104
        
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirHerreria_Err
        

        Dim i As Long

100     For i = 1 To UBound(ArmasHerrero)

102         If ArmasHerrero(i) = ItemIndex Then
104             PuedeConstruirHerreria = True
                Exit Function

            End If

106     Next i

108     For i = 1 To UBound(ArmadurasHerrero)

110         If ArmadurasHerrero(i) = ItemIndex Then
112             PuedeConstruirHerreria = True
                Exit Function

            End If

114     Next i

116     PuedeConstruirHerreria = False

        
        Exit Function

PuedeConstruirHerreria_Err:
118     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirHerreria", Erl)
120
        
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
        
102     If UserList(UserIndex).flags.Privilegios And (PlayerType.Consejero) Then
            Exit Sub
        End If
        
104     If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
106         Call HerreroQuitarMateriales(UserIndex, ItemIndex)
108         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
110         Call WriteUpdateSta(UserIndex)
            ' AGREGAR FX
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 253, 25, False, ObjData(ItemIndex).GrhIndex))
112         If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el arma!", FontTypeNames.FONTTYPE_INFO)
114             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
116         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otEscudo Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el escudo!", FontTypeNames.FONTTYPE_INFO)
118             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
120         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCasco Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el casco!", FontTypeNames.FONTTYPE_INFO)
122             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
124         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
                'Call WriteConsoleMsg(UserIndex, "Has construido la armadura!", FontTypeNames.FONTTYPE_INFO)
126             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)

            End If

            Dim MiObj As obj

128         MiObj.amount = 1
130         MiObj.ObjIndex = ItemIndex

132         If Not MeterItemEnInventario(UserIndex, MiObj) Then
134             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            'End If
    
136         Call SubirSkill(UserIndex, eSkill.Herreria)
138         Call UpdateUserInv(True, UserIndex, 0)
140         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

142         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

HerreroConstruirItem_Err:
144     Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroConstruirItem", Erl)
146
        
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirCarpintero_Err
        

        Dim i As Long

100     For i = 1 To UBound(ObjCarpintero)

102         If ObjCarpintero(i) = ItemIndex Then
104             PuedeConstruirCarpintero = True
                Exit Function

            End If

106     Next i

108     PuedeConstruirCarpintero = False

        
        Exit Function

PuedeConstruirCarpintero_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirCarpintero", Erl)
112
        
End Function

Public Function PuedeConstruirAlquimista(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirAlquimista_Err
        

        Dim i As Long

100     For i = 1 To UBound(ObjAlquimista)

102         If ObjAlquimista(i) = ItemIndex Then
104             PuedeConstruirAlquimista = True
                Exit Function

            End If

106     Next i

108     PuedeConstruirAlquimista = False

        
        Exit Function

PuedeConstruirAlquimista_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirAlquimista", Erl)
112
        
End Function

Public Function PuedeConstruirSastre(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirSastre_Err
        

        Dim i As Long

100     For i = 1 To UBound(ObjSastre)

102         If ObjSastre(i) = ItemIndex Then
104             PuedeConstruirSastre = True
                Exit Function

            End If

106     Next i

108     PuedeConstruirSastre = False

        
        Exit Function

PuedeConstruirSastre_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirSastre", Erl)
112
        
End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub

102     If UserList(UserIndex).flags.Privilegios And (PlayerType.Consejero) Then
            Exit Sub
        End If
        
104     If ItemIndex = 0 Then Exit Sub

106     If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
            ' Antes de usar la herramienta deberias equipartela.
108         Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
110         Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        
112     If CarpinteroTieneMateriales(UserIndex, ItemIndex) _
                And UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria _
                And PuedeConstruirCarpintero(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 5 Then
    
114         If UserList(UserIndex).Stats.MinSta > 2 Then
116             Call QuitarSta(UserIndex, 2)
        
            Else
118             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para trabajar.", FontTypeNames.FONTTYPE_INFO)
120             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If
    
122         Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido un objeto!", FontTypeNames.FONTTYPE_INFO)
            'Call WriteOroOverHead(UserIndex, 1, UserList(UserIndex).Char.CharIndex)
124         Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
    
            Dim MiObj As obj

126         MiObj.amount = 1
128         MiObj.ObjIndex = ItemIndex
             ' AGREGAR FX
130         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
132         If Not MeterItemEnInventario(UserIndex, MiObj) Then
134             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            ' End If
    
136         Call SubirSkill(UserIndex, eSkill.Carpinteria)
            'Call UpdateUserInv(True, UserIndex, 0)
138         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

140         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

CarpinteroConstruirItem_Err:
142     Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroConstruirItem", Erl)

        
End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaConstruirItem_Err
        

        Rem Debug.Print UserList(UserIndex).Invent.HerramientaEqpObjIndex

100     If Not UserList(UserIndex).Stats.MinSta > 0 Then
102         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If AlquimistaTieneMateriales(UserIndex, ItemIndex) _
                And UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) >= ObjData(ItemIndex).SkPociones _
                And PuedeConstruirAlquimista(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 4 Then
        
106         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 25
108         Call WriteUpdateSta(UserIndex)
            
            ' AGREGAR FX
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 253, 25, False, ObjData(ItemIndex).GrhIndex))
110         Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
            Dim MiObj As obj

114         MiObj.amount = 1
116         MiObj.ObjIndex = ItemIndex

118         If Not MeterItemEnInventario(UserIndex, MiObj) Then
120             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ''If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            'End If
    
122         Call SubirSkill(UserIndex, eSkill.Alquimia)
124         Call UpdateUserInv(True, UserIndex, 0)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

126         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

AlquimistaConstruirItem_Err:
128     Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaConstruirItem", Erl)
130
        
End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub

102     If Not UserList(UserIndex).Stats.MinSta > 0 Then
104         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

106     If SastreTieneMateriales(UserIndex, ItemIndex) _
                And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkMAGOria _
                And PuedeConstruirSastre(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 9 Then
        
108         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        
110         Call WriteUpdateSta(UserIndex)
    
112         Call SastreQuitarMateriales(UserIndex, ItemIndex)
    
            ' If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
            ' Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
114         Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
    '112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5))
            ' UserList(UserIndex).flags.UltimoMensaje = 9
            ' End If
        
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(63, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
            Dim MiObj As obj

118         MiObj.amount = 1
120         MiObj.ObjIndex = ItemIndex

122         If Not MeterItemEnInventario(UserIndex, MiObj) Then
124             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
    
126         Call SubirSkill(UserIndex, eSkill.Herreria)
128         Call UpdateUserInv(True, UserIndex, 0)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

130         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If
    
        
        Exit Sub

SastreConstruirItem_Err:
132     Call TraceError(Err.Number, Err.Description, "Trabajo.SastreConstruirItem", Erl)
134
        
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales, ByVal cant As Byte) As Integer
        
        On Error GoTo MineralesParaLingote_Err
        

100     Select Case Lingote

            Case iMinerales.HierroCrudo
102             MineralesParaLingote = 13 * cant

104         Case iMinerales.PlataCruda
106             MineralesParaLingote = 25 * cant

108         Case iMinerales.OroCrudo
110             MineralesParaLingote = 50 * cant

112         Case Else
114             MineralesParaLingote = 10000

        End Select

        
        Exit Function

MineralesParaLingote_Err:
116     Call TraceError(Err.Number, Err.Description, "Trabajo.MineralesParaLingote", Erl)
118
        
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
            On Error GoTo DoLingotes_Err

            Dim Slot As Integer
            Dim obji As Integer
            Dim cant As Byte
            Dim necesarios As Integer

100         If UserList(UserIndex).Stats.MinSta > 2 Then
102             Call QuitarSta(UserIndex, 2)

            Else
104             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
106             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If

108         Slot = UserList(UserIndex).flags.TargetObjInvSlot
110         obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex

112         cant = RandomNumber(10, 20)
114         necesarios = MineralesParaLingote(obji, cant)

116         If UserList(UserIndex).Invent.Object(Slot).amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
118             Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
120             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If

122         UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount - MineralesParaLingote(obji, cant)

124         If UserList(UserIndex).Invent.Object(Slot).amount < 1 Then
126             UserList(UserIndex).Invent.Object(Slot).amount = 0
128             UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0

            End If

            Dim nPos  As WorldPos

            Dim MiObj As obj

130         MiObj.amount = cant
132         MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex

134         If Not MeterItemEnInventario(UserIndex, MiObj) Then
136             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If

138         Call UpdateUserInv(False, UserIndex, Slot)
140         Call WriteTextCharDrop(UserIndex, "+" & cant, UserList(UserIndex).Char.CharIndex, vbWhite)
142         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(41, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
144         Call SubirSkill(UserIndex, eSkill.Mineria)

146         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

148         If UserList(UserIndex).Counters.Trabajando = 1 And Not UserList(UserIndex).flags.UsandoMacro Then
150             Call WriteMacroTrabajoToggle(UserIndex, True)

            End If

            Exit Sub

DoLingotes_Err:
152         Call TraceError(Err.Number, Err.Description, "Trabajo.DoLingotes", Erl)
154

End Sub

Function ModAlquimia(ByVal clase As eClass) As Integer
        On Error GoTo ModAlquimia_Err

100     Select Case clase

            Case eClass.Druid
102             ModAlquimia = 1

104         Case eClass.Trabajador
106             ModAlquimia = 1

108         Case Else
110             ModAlquimia = 3

        End Select

        Exit Function

ModAlquimia_Err:
112     Call TraceError(Err.Number, Err.Description, "Trabajo.ModAlquimia", Erl)
114

End Function

Function ModSastre(ByVal clase As eClass) As Integer
        
        On Error GoTo ModSastre_Err
        

100     Select Case clase

            Case eClass.Trabajador
102             ModSastre = 1

104         Case Else
106             ModSastre = 3

        End Select

        
        Exit Function

ModSastre_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ModSastre", Erl)
110
        
End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer
        
        On Error GoTo ModCarpinteria_Err
        

100     Select Case clase

            Case eClass.Trabajador
102             ModCarpinteria = 1

104         Case Else
106             ModCarpinteria = 3

        End Select

        
        Exit Function

ModCarpinteria_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ModCarpinteria", Erl)
110
        
End Function

Function ModHerreria(ByVal clase As eClass) As Single
        
        On Error GoTo ModHerreriA_Err
        

100     Select Case clase

            Case eClass.Trabajador
102             ModHerreria = 1

104         Case Else
106             ModHerreria = 3

        End Select

        
        Exit Function

ModHerreriA_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ModHerreriA", Erl)
110
        
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
        
        On Error GoTo DoAdminInvisible_Err
    
100     With UserList(UserIndex)
    
102         If .flags.AdminInvisible = 0 Then
                
104             .flags.AdminInvisible = 1
106             .flags.invisible = 1
108             .flags.Oculto = 1
            
                '.flags.OldBody = .Char.Body
                '.flags.OldHead = .Char.Head
            
                '.Char.Body = 0
                '.Char.Head = 0
            
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            
112             Call SendData(SendTarget.ToPCAreaButGMs, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex, True))
            
            Else
        
114             .flags.AdminInvisible = 0
116             .flags.invisible = 0
118             .flags.Oculto = 0
120             .Counters.TiempoOculto = 0
            
                '.Char.Body = .flags.OldBody
                '.Char.Head = .flags.OldHead
            
122             Call MakeUserChar(True, 0, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, 1)
124             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            
            End If
        
            'Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    
        End With

        Exit Sub

DoAdminInvisible_Err:
126     Call TraceError(Err.Number, Err.Description, "Trabajo.DoAdminInvisible", Erl)

128
        
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo TratarDeHacerFogata_Err
        

        Dim Suerte    As Byte

        Dim exito     As Byte

        Dim obj       As obj

        Dim posMadera As WorldPos

100     If Not LegalPos(Map, X, Y) Then Exit Sub

102     With posMadera
104         .Map = Map
106         .X = X
108         .Y = Y

        End With

110     If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
112         Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
116         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

118     If UserList(UserIndex).flags.Muerto = 1 Then
120         Call WriteConsoleMsg(UserIndex, "No podés hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

122     If MapData(Map, X, Y).ObjInfo.amount < 3 Then
124         Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

126     If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
128         Suerte = 3
130     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
132         Suerte = 2
134     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
136         Suerte = 1

        End If

138     exito = RandomNumber(1, Suerte)

140     If exito = 1 Then
142         obj.ObjIndex = FOGATA_APAG
144         obj.amount = MapData(Map, X, Y).ObjInfo.amount \ 3
    
146         Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.amount & " ramitas.", FontTypeNames.FONTTYPE_INFO)
    
148         Call MakeObj(obj, Map, X, Y)
    
            'Seteamos la fogata como el nuevo TargetObj del user
150         UserList(UserIndex).flags.TargetObj = FOGATA_APAG

        End If

152     Call SubirSkill(UserIndex, Supervivencia)

        
        Exit Sub

TratarDeHacerFogata_Err:
154     Call TraceError(Err.Number, Err.Description, "Trabajo.TratarDeHacerFogata", Erl)
156
        
End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal RedDePesca As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte       As Integer
        Dim res          As Long
        Dim RestaStamina As Byte

100     RestaStamina = IIf(RedDePesca, 5, 1)
    
102     With UserList(UserIndex)

104         If .flags.Privilegios And (PlayerType.Consejero) Then
                Exit Sub
            End If
            
106         If .Stats.MinSta > RestaStamina Then
108             Call QuitarSta(UserIndex, RestaStamina)
        
            Else
            
110             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para pescar.", FontTypeNames.FONTTYPE_INFO)
            
112             Call WriteMacroTrabajoToggle(UserIndex, False)
            
                Exit Sub

            End If

            Dim Skill As Integer

114         Skill = .Stats.UserSkills(eSkill.Pescar)
        
116         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
118         res = RandomNumber(1, Suerte)
            'HarThaoS: Movimiento de caña, lo saco. Se hace exponencial la cantidad de paquetes dependiendo la cantida de usuarios
'120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))

122         If res < 6 Then

                Dim nPos  As WorldPos
                Dim MiObj As obj

124             MiObj.amount = IIf(.clase = Trabajador, RandomNumber(1, 3), 1) * RecoleccionMult
126             MiObj.ObjIndex = ObtenerPezRandom(ObjData(.Invent.HerramientaEqpObjIndex).Power)
                 ' AGREGAR FX
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
128             If MiObj.ObjIndex = 0 Then Exit Sub
        
130             If Not MeterItemEnInventario(UserIndex, MiObj) Then
132                 Call TirarItemAlPiso(.Pos, MiObj)
                End If

134             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)
                 
                 If MapInfo(.Pos.Map).Seguro = 1 Then
302                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
301                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                End If
                ' Al pescar también podés sacar cosas raras (se setean desde RecursosEspeciales.dat)
                Dim i As Integer

                ' Por cada drop posible
136             For i = 1 To UBound(EspecialesPesca)
                    ' Tiramos al azar entre 1 y la probabilidad
138                 res = RandomNumber(1, IIf(RedDePesca, EspecialesPesca(i).Data * 2, EspecialesPesca(i).Data)) ' Red de pesca chance x2 (revisar)
            
                    ' Si tiene suerte y le pega
140                 If res = 1 Then
142                     MiObj.ObjIndex = EspecialesPesca(i).ObjIndex
144                     MiObj.amount = 1 ' Solo un item por vez
                
146                     If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                    
                        ' Le mandamos un mensaje
148                     Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(EspecialesPesca(i).ObjIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
                    End If

                Next

            Else
                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, GRH_FALLO_PESCA))
            End If
    
150         Call SubirSkill(UserIndex, eSkill.Pescar)
    
152         .Counters.Trabajando = .Counters.Trabajando + 1
    
            'Ladder 06/07/14 Activamos el macro de trabajo
154         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
156             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
158     Call LogError("Error en DoPescar. Error " & Err.Number & " - " & Err.Description)

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadronIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadronIndex As Integer, ByVal VictimaIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 05/04/2010
        'Last Modification By: ZaMa
        '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadronIndex) when the thief stoles gold. (MarKoxX)
        '27/11/2009: ZaMa - Optimizacion de codigo.
        '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
        '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
        '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
        '23/04/2010: ZaMa - No se puede robar mas sin energia.
        '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
        '*************************************************

        On Error GoTo ErrHandler

        Dim OtroUserIndex As Integer
        
100     If UserList(LadronIndex).flags.Privilegios And (PlayerType.Consejero) Then Exit Sub

102     If MapInfo(UserList(VictimaIndex).Pos.Map).Seguro = 1 Then Exit Sub
    
104     If UserList(VictimaIndex).flags.EnConsulta Then
106         Call WriteConsoleMsg(LadronIndex, "¡No puedes robar a usuarios en consulta!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
108     With UserList(LadronIndex)
    
110         If .flags.Seguro Then
        
112             If Status(VictimaIndex) = Ciudadano Then
114                 Call WriteConsoleMsg(LadronIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            Else

116             If esArmada(LadronIndex) Then
            
118                 If Status(VictimaIndex) = Ciudadano Or esArmada(VictimaIndex) Then
120                     Call WriteConsoleMsg(LadronIndex, "Los miembros del Ejército Real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub

                    End If

                End If

            End If
        
            ' Caos robando a caos?
122         If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
124             Call WriteConsoleMsg(LadronIndex, "No puedes robar a otros miembros de la Legión Oscura.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
        
126         If TriggerZonaPelea(LadronIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
            ' Tiene energia?
128         If .Stats.MinSta < 15 Then
        
130             If .genero = eGenero.Hombre Then
132                 Call WriteConsoleMsg(LadronIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
                
                Else
134                 Call WriteConsoleMsg(LadronIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
                Exit Sub

            End If
        
136         If .GuildIndex > 0 Then
        
138             If .flags.SeguroClan Then
            
140                 If .GuildIndex = UserList(VictimaIndex).GuildIndex Then
142                     Call WriteConsoleMsg(LadronIndex, "No podes robarle a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

                End If

            End If

            ' Quito energia
144         Call QuitarSta(LadronIndex, 15)

146         If UserList(VictimaIndex).flags.Privilegios And PlayerType.user Then

                Dim Probabilidad As Byte

                Dim res          As Integer

                Dim RobarSkill   As Byte
            
148             RobarSkill = .Stats.UserSkills(eSkill.Robar)

150             If (RobarSkill > 0 And RobarSkill < 10) Then
152                 Probabilidad = 1
154             ElseIf (RobarSkill >= 10 And RobarSkill <= 20) Then
156                 Probabilidad = 5
158             ElseIf (RobarSkill >= 20 And RobarSkill <= 30) Then
160                 Probabilidad = 10
162             ElseIf (RobarSkill >= 30 And RobarSkill <= 40) Then
164                 Probabilidad = 15
166             ElseIf (RobarSkill >= 40 And RobarSkill <= 50) Then
168                 Probabilidad = 25
170             ElseIf (RobarSkill >= 50 And RobarSkill <= 60) Then
172                 Probabilidad = 35
174             ElseIf (RobarSkill >= 60 And RobarSkill <= 70) Then
176                 Probabilidad = 40
178             ElseIf (RobarSkill >= 70 And RobarSkill <= 80) Then
180                 Probabilidad = 55
182             ElseIf (RobarSkill >= 80 And RobarSkill <= 90) Then
184                 Probabilidad = 70
186             ElseIf (RobarSkill >= 90 And RobarSkill < 100) Then
188                 Probabilidad = 80
190             ElseIf (RobarSkill = 100) Then
192                 Probabilidad = 90
                End If

194             If (RandomNumber(1, 100) < Probabilidad) Then 'Exito robo
                
196                 If UserList(VictimaIndex).flags.Comerciando Then
198                     OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
200                     If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
202                         Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, ¡te están robando!", FontTypeNames.FONTTYPE_TALK)
204                         Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado, al otro usuario le robaron.", FontTypeNames.FONTTYPE_TALK)
                        
206                         Call LimpiarComercioSeguro(VictimaIndex)

                        End If

                    End If
               
208                 If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then '50% de robar items
                    
210                     If TieneObjetosRobables(VictimaIndex) Then
212                         Call RobarObjeto(LadronIndex, VictimaIndex)
                        Else
214                         Call WriteConsoleMsg(LadronIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else '50% de robar oro

216                     If UserList(VictimaIndex).Stats.GLD > 0 Then

                            Dim n     As Long

                            Dim Extra As Single
                            
                            ' Multiplicador extra por niveles
218                         If (.Stats.ELV < 25) Then
220                             Extra = 1
222                         ElseIf (.Stats.ELV < 35) Then
224                             Extra = 1.05
226                         ElseIf (.Stats.ELV >= 35 And .Stats.ELV <= 40) Then
228                             Extra = 1.1
230                         ElseIf (.Stats.ELV >= 41 And .Stats.ELV < 45) Then
232                             Extra = 1.15
234                         ElseIf (.Stats.ELV >= 45 And .Stats.ELV <= 47) Then
236                             Extra = 1.2

                            End If
                            
238                         If .clase = eClass.Thief Then

                                'Si no tiene puestos los guantes de hurto roba un 50% menos.
240                             If .Invent.NudilloObjIndex > 0 Then
242                                 If ObjData(.Invent.NudilloObjIndex).Subtipo = 5 Then
244                                     n = RandomNumber(.Stats.ELV * 50 * Extra, .Stats.ELV * 100 * Extra) * OroMult
                                    Else
246                                     n = RandomNumber(.Stats.ELV * 25 * Extra, .Stats.ELV * 50 * Extra) * OroMult

                                    End If

                                Else
248                                 n = RandomNumber(.Stats.ELV * 25 * Extra, .Stats.ELV * 50 * Extra) * OroMult

                                End If
    
                            Else
250                             n = RandomNumber(1, 100) * OroMult
    
                            End If

252                         If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                        
254                         UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                        
256                         .Stats.GLD = .Stats.GLD + n

258                         If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO
                        
260                         Call WriteConsoleMsg(LadronIndex, "Le has robado " & PonerPuntos(n) & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_New_Rojo_Salmon)
262                         Call WriteConsoleMsg(VictimaIndex, UserList(LadronIndex).Name & " te ha robado " & PonerPuntos(n) & " monedas de oro.", FontTypeNames.FONTTYPE_New_Rojo_Salmon)
264                         Call WriteUpdateGold(LadronIndex) 'Le actualizamos la billetera al ladron
                        
266                         Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Else
268                         Call WriteConsoleMsg(LadronIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                
270                 Call SubirSkill(LadronIndex, eSkill.Robar)
            
                Else
272                 Call WriteConsoleMsg(LadronIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
274                 Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
276                 Call SubirSkill(LadronIndex, eSkill.Robar)

                End If

278             If Status(LadronIndex) = Ciudadano Then Call VolverCriminal(LadronIndex)

            End If

        End With

        Exit Sub

ErrHandler:
282     Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)

End Sub

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
        ' Agregué los barcos
        ' Agrego poción negra
        ' Esta funcion determina qué objetos son robables.
        
        On Error GoTo ObjEsRobable_Err
        

        Dim OI As Integer

100     OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

102     ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).donador = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otRunas And ObjData(OI).Instransferible = 0 And ObjData(OI).OBJType <> eOBJType.otMonturas And Not (ObjData(OI).OBJType = eOBJType.otPociones And ObjData(OI).TipoPocion = 21)

        
        Exit Function

ObjEsRobable_Err:
104     Call TraceError(Err.Number, Err.Description, "Trabajo.ObjEsRobable", Erl)
106
        
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Private Sub RobarObjeto(ByVal LadronIndex As Integer, ByVal VictimaIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
        '***************************************************
        
        On Error GoTo RobarObjeto_Err
    
        

        Dim flag As Boolean
        Dim i    As Integer

100     flag = False

102     With UserList(VictimaIndex)

104         If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final del inventario?
106             i = 1

108             Do While Not flag And i <= .CurrentInventorySlots

                    'Hay objeto en este slot?
110                 If .Invent.Object(i).ObjIndex > 0 Then
                
112                     If ObjEsRobable(VictimaIndex, i) Then
                    
114                         If RandomNumber(1, 10) < 4 Then flag = True
                        
                        End If

                    End If

116                 If Not flag Then i = i + 1
                Loop
            Else
118             i = .CurrentInventorySlots

120             Do While Not flag And i > 0

                    'Hay objeto en este slot?
122                 If .Invent.Object(i).ObjIndex > 0 Then
                
124                     If ObjEsRobable(VictimaIndex, i) Then
                    
126                         If RandomNumber(1, 10) < 4 Then flag = True
                        
                        End If

                    End If

128                 If Not flag Then i = i - 1
                Loop

            End If
    
130         If flag Then

                Dim MiObj     As obj
                Dim num       As Integer
                Dim ObjAmount As Integer

132             ObjAmount = .Invent.Object(i).amount

                'Cantidad al azar entre el 3 y el 6% del total, con minimo 1.
134             num = MaximoInt(1, RandomNumber(ObjAmount * 0.03, ObjAmount * 0.06))

136             MiObj.amount = num
138             MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
140             .Invent.Object(i).amount = ObjAmount - num
                    
142             If .Invent.Object(i).amount <= 0 Then
144                 Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

                End If
                
146             Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
148             If Not MeterItemEnInventario(LadronIndex, MiObj) Then
150                 Call TirarItemAlPiso(UserList(LadronIndex).Pos, MiObj)
                
                End If
        
152             If UserList(LadronIndex).clase = eClass.Thief Then
154                 Call WriteConsoleMsg(LadronIndex, "Has robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_New_Rojo_Salmon)
156                 Call WriteConsoleMsg(VictimaIndex, UserList(LadronIndex).Name & " te ha robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_New_Rojo_Salmon)
                Else
158                 Call WriteConsoleMsg(LadronIndex, "Has hurtado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_New_Rojo_Salmon)
                
                End If

            Else
160             Call WriteConsoleMsg(LadronIndex, "No has logrado robar ningun objeto.", FontTypeNames.FONTTYPE_INFO)

            End If

            'If exiting, cancel de quien es robado
162         Call CancelExit(VictimaIndex)

        End With

        
        Exit Sub

RobarObjeto_Err:
164     Call TraceError(Err.Number, Err.Description, "Trabajo.RobarObjeto", Erl)

        
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarSta_Err
        
100     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

102     If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
104     If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub
106     Call WriteUpdateSta(UserIndex)

        Exit Sub

QuitarSta_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.QuitarSta", Erl)
110
        
End Sub

Public Sub DoRaices(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
    
100     With UserList(UserIndex)
    
102         If .flags.Privilegios And (PlayerType.Consejero) Then
                Exit Sub
            End If
            
104         If .Stats.MinSta > 2 Then
106             Call QuitarSta(UserIndex, 2)
        
            Else
            
108             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para obtener raices.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            Dim Skill As Integer
112             Skill = .Stats.UserSkills(eSkill.Alquimia)
        
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
116         res = RandomNumber(1, Suerte)
    
'118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
    
            Rem Ladder 06/08/14 Subo un poco la probabilidad de sacar raices... porque era muy lento
120         If res < 7 Then
    
                Dim nPos  As WorldPos
                Dim MiObj As obj
        
                'If .clase = eClass.Druid Then
                'MiObj.Amount = RandomNumber(6, 8)
                ' Else
122             MiObj.amount = RandomNumber(5, 7)
                ' End If
       
124             If ObjData(.Invent.HerramientaEqpObjIndex).donador = 1 Then
126                 MiObj.amount = MiObj.amount * 2
                End If
       
128             MiObj.amount = MiObj.amount * RecoleccionMult
130             MiObj.ObjIndex = Raices
        
132             MapData(.Pos.Map, X, Y).ObjInfo.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount - MiObj.amount
    
134             If MapData(.Pos.Map, X, Y).ObjInfo.amount < 0 Then
136                 MapData(.Pos.Map, X, Y).ObjInfo.amount = 0
    
138                 Call AgregarItemLimpieza(.Pos.Map, X, Y)
                
                End If
        
140             If Not MeterItemEnInventario(UserIndex, MiObj) Then
            
142                 Call TirarItemAlPiso(.Pos, MiObj)
            
                End If
        
                'Call WriteConsoleMsg(UserIndex, "¡Has conseguido algunas raices!", FontTypeNames.FONTTYPE_INFO)
144             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)
146             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(60, .Pos.X, .Pos.Y))
            Else
148             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(61, .Pos.X, .Pos.Y))
    
            End If
    
150         Call SubirSkill(UserIndex, eSkill.Alquimia)
    
152         .Counters.Trabajando = .Counters.Trabajando + 1
    
154         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
156             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
158     Call LogError("Error en DoRaices")

End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)
        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer

100     With UserList(UserIndex)

102          If .flags.Privilegios And (PlayerType.Consejero) Then
                Exit Sub
             End If


                'EsfuerzoTalarLeñador = 1
104         If .Stats.MinSta > 2 Then
106             Call QuitarSta(UserIndex, 2)

            Else
108             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para talar.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If

            Dim Skill As Integer

112         Skill = .Stats.UserSkills(eSkill.Talar)
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

116         res = RandomNumber(1, Suerte)
'118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))

120         If res < 6 Then

                Dim nPos  As WorldPos

                Dim MiObj As obj

122             Call ActualizarRecurso(.Pos.Map, X, Y)
124             MapData(.Pos.Map, X, Y).ObjInfo.Data = GetTickCount() ' Ultimo uso
    
126             MiObj.amount = IIf(.clase = Trabajador, 5, RandomNumber(1, 2)) * RecoleccionMult

128             If ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex).Elfico = 0 Then
130                 MiObj.ObjIndex = Leña
                Else
132                 MiObj.ObjIndex = LeñaElfica
                End If


134             If MiObj.amount > MapData(.Pos.Map, X, Y).ObjInfo.amount Then
136                 MiObj.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount
                End If
            
138             MapData(.Pos.Map, X, Y).ObjInfo.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount - MiObj.amount
                 ' AGREGAR FX
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
140             If Not MeterItemEnInventario(UserIndex, MiObj) Then
142                 Call TirarItemAlPiso(.Pos, MiObj)
                End If
    
144             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)

                If MapInfo(.Pos.Map).Seguro = 1 Then
146                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                Else
145                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                End If

                ' Al talar también podés dropear cosas raras (se setean desde RecursosEspeciales.dat)
                Dim i As Integer

                ' Por cada drop posible
148             For i = 1 To UBound(EspecialesTala)
                    ' Tiramos al azar entre 1 y la probabilidad
150                 res = RandomNumber(1, EspecialesTala(i).Data)

                    ' Si tiene suerte y le pega
152                 If res = 1 Then
154                     MiObj.ObjIndex = EspecialesTala(i).ObjIndex
156                     MiObj.amount = 1 ' Solo un item por vez

                        ' Tiro siempre el item al piso, me parece más rolero, como que cae del árbol :P
158                     Call TirarItemAlPiso(.Pos, MiObj)
                    End If

160             Next i

            Else
162             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(64, .Pos.X, .Pos.Y))

            End If
        
164         Call SubirSkill(UserIndex, eSkill.Talar)
        
166         .Counters.Trabajando = .Counters.Trabajando + 1
    
168         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
170             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With

        Exit Sub

ErrHandler:
172     Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte      As Integer
        Dim res         As Integer
        Dim Metal       As Integer
        Dim Yacimiento  As ObjData

100     With UserList(UserIndex)
    
102         If .flags.Privilegios And (PlayerType.Consejero) Then
                Exit Sub
            End If
    
            'Por Ladder 06/07/2014 Cuando la estamina llega a 0 , el macro se desactiva
104         If .Stats.MinSta > 2 Then
106             Call QuitarSta(UserIndex, 2)
            Else
108             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            'Por Ladder 06/07/2014
    
            Dim Skill As Integer
    
112         Skill = .Stats.UserSkills(eSkill.Mineria)
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
116         res = RandomNumber(1, Suerte)
        
'118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        
120         If res <= 5 Then
    
                Dim MiObj As obj
                Dim nPos  As WorldPos
            
122             Call ActualizarRecurso(.Pos.Map, X, Y)
124             MapData(.Pos.Map, X, Y).ObjInfo.Data = GetTickCount() ' Ultimo uso

126             Yacimiento = ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex)
            
128             MiObj.ObjIndex = Yacimiento.MineralIndex
130             MiObj.amount = IIf(.clase = Trabajador, 5, RandomNumber(1, 2)) * RecoleccionMult
            
132             If MiObj.amount > MapData(.Pos.Map, X, Y).ObjInfo.amount Then
134                 MiObj.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount
                End If
            
136             MapData(.Pos.Map, X, Y).ObjInfo.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount - MiObj.amount
        
138             If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                 ' AGREGAR FX
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
                
139             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)

140             Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
                If MapInfo(.Pos.Map).Seguro = 1 Then
141                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(15, .Pos.X, .Pos.Y))
                Else
142                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(15, .Pos.X, .Pos.Y))
                End If
            
                ' Al minar también puede dropear una gema
                Dim i As Integer
    
                ' Por cada drop posible
144             For i = 1 To Yacimiento.CantItem
                    ' Tiramos al azar entre 1 y la probabilidad
146                 res = RandomNumber(1, Yacimiento.Item(i).amount)
                
                    ' Si tiene suerte y le pega
148                 If res = 1 Then
                        ' Se lo metemos al inventario (o lo tiramos al piso)
150                     MiObj.ObjIndex = Yacimiento.Item(i).ObjIndex
152                     MiObj.amount = 1 ' Solo una gema por vez
                    
154                     If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)

                        ' Le mandamos un mensaje
156                     Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(Yacimiento.Item(i).ObjIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
                    End If
    
                Next
            
            Else
158             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(2185, .Pos.X, .Pos.Y))

            End If

160         Call SubirSkill(UserIndex, eSkill.Mineria)

162         .Counters.Trabajando = .Counters.Trabajando + 1

164         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
166             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
168     Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
        
        On Error GoTo DoMeditar_Err

        Dim Mana As Long
        
100     With UserList(UserIndex)

102         .Counters.TimerMeditar = .Counters.TimerMeditar + 1

104         If .Counters.TimerMeditar >= IntervaloMeditar Then

106             Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, 50 + .Stats.UserSkills(eSkill.Meditar) * 0.5))

108             If Mana <= 0 Then Mana = 1

110             If .Stats.MinMAN + Mana >= .Stats.MaxMAN Then

112                 .Stats.MinMAN = .Stats.MaxMAN
114                 .flags.Meditando = False
116                 .Char.FX = 0
                    
118                 Call WriteUpdateMana(UserIndex)
120                 Call SubirSkill(UserIndex, Meditar)

122                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
                
                Else
                    
124                 .Stats.MinMAN = .Stats.MinMAN + Mana
                    
126                 Call WriteUpdateMana(UserIndex)
128                 Call SubirSkill(UserIndex, Meditar)

                End If

130             .Counters.TimerMeditar = 0
            End If

        End With

        
        Exit Sub

DoMeditar_Err:
132     Call TraceError(Err.Number, Err.Description, "Trabajo.DoMeditar", Erl)
134
        
End Sub

Public Sub DoMontar(ByVal UserIndex As Integer, ByRef Montura As ObjData, ByVal Slot As Integer)
        On Error GoTo DoMontar_Err

100     With UserList(UserIndex)
102         If PuedeUsarObjeto(UserIndex, .Invent.Object(Slot).ObjIndex, True) > 0 Then
                Exit Sub
            End If

104         If .flags.Montado = 0 And .Counters.EnCombate > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Estás en combate, debes aguardar " & .Counters.EnCombate & " segundo(s) para montar...", FontTypeNames.FONTTYPE_INFOBOLD)
                Exit Sub
            End If

108         If .flags.EnReto Then
110             Call WriteConsoleMsg(UserIndex, "No podés montar en un reto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

112         If (.flags.Oculto = 1 Or .flags.invisible = 1) And .flags.AdminInvisible = 0 Then
114             Call WriteConsoleMsg(UserIndex, "No podés montar estando oculto o invisible.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
                Call WriteConsoleMsg(UserIndex, "Pierdes el efecto del mimetismo.", FontTypeNames.FONTTYPE_INFO)
                .Counters.Mimetismo = 0
                .flags.Mimetizado = e_EstadoMimetismo.Desactivado
                Call RefreshCharStatus(UserIndex)
            End If

116         If .flags.Montado = 0 And (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger > 10) Then
118             Call WriteConsoleMsg(UserIndex, "No podés montar aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

120         If .flags.Meditando Then
122             .flags.Meditando = False
124             .Char.FX = 0
126             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
            End If

128         If .flags.Montado = 1 And .Invent.MonturaObjIndex > 0 Then
130             If ObjData(.Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
132                 Call UpdateUserInv(False, UserIndex, .Invent.MonturaSlot)
                End If

            End If

134         .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
136         .Invent.MonturaSlot = Slot

138         If .flags.Montado = 0 Then
140             .Char.Body = Montura.Ropaje
142             .Char.Head = .OrigChar.Head
144             .Char.ShieldAnim = NingunEscudo
146             .Char.WeaponAnim = NingunArma
148             .Char.CascoAnim = .Char.CascoAnim
150             .flags.Montado = 1
            Else
152             .flags.Montado = 0
154             .Char.Head = .OrigChar.Head

156             If .Invent.ArmourEqpObjIndex > 0 Then
158                 .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje

                Else
160                 Call DarCuerpoDesnudo(UserIndex)

                End If

162             If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

164             If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

166             If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim

            End If

168         Call ActualizarVelocidadDeUsuario(UserIndex)
170         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

172         Call UpdateUserInv(False, UserIndex, Slot)
174         Call WriteEquiteToggle(UserIndex)
        End With

        Exit Sub

DoMontar_Err:
176     Call TraceError(Err.Number, Err.Description, "Trabajo.DoMontar", Erl)
178

End Sub

Public Sub ActualizarRecurso(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo ActualizarRecurso_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex

        Dim TiempoActual As Long

102     TiempoActual = GetTickCount()

        ' Data = Ultimo uso
104     If (TiempoActual - MapData(Map, X, Y).ObjInfo.Data) * 0.001 > ObjData(ObjIndex).TiempoRegenerar Then
106         MapData(Map, X, Y).ObjInfo.amount = ObjData(ObjIndex).VidaUtil
108         MapData(Map, X, Y).ObjInfo.Data = &H7FFFFFFF   ' Ultimo uso = Max Long

        End If

        
        Exit Sub

ActualizarRecurso_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.ActualizarRecurso", Erl)
112
        
End Sub

Public Function ObtenerPezRandom(ByVal PoderCania As Integer) As Long
        
        On Error GoTo ObtenerPezRandom_Err
        

        Dim i As Long, SumaPesos As Long, ValorGenerado As Long
    
100     If PoderCania > UBound(PesoPeces) Then PoderCania = UBound(PesoPeces)
102     SumaPesos = PesoPeces(PoderCania)

104     ValorGenerado = RandomNumber(0, SumaPesos - 1)

106     ObtenerPezRandom = Peces(BinarySearchPeces(ValorGenerado)).ObjIndex

        Exit Function

ObtenerPezRandom_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ObtenerPezRandom", Erl)
110
        
End Function

Function ModDomar(ByVal clase As eClass) As Integer
        
        On Error GoTo ModDomar_Err
    
        

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
100     Select Case clase

            Case eClass.Druid
102             ModDomar = 6

104         Case eClass.Hunter
106             ModDomar = 6

108         Case eClass.Cleric
110             ModDomar = 7

112         Case Else
114             ModDomar = 10

        End Select

        
        Exit Function

ModDomar_Err:
116     Call TraceError(Err.Number, Err.Description, "Trabajo.ModDomar", Erl)

        
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
        
        On Error GoTo FreeMascotaIndex_Err
    
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 02/03/09
        '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
        '***************************************************
        Dim j As Integer

100     For j = 1 To MAXMASCOTAS

102         If UserList(UserIndex).MascotasType(j) = 0 Then
104             FreeMascotaIndex = j
                Exit Function

            End If

106     Next j

        
        Exit Function

FreeMascotaIndex_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.FreeMascotaIndex", Erl)

        
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Nacho (Integer)
        'Last Modification: 01/05/2010
        '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
        '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
        '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
        '***************************************************

        On Error GoTo ErrHandler

        Dim puntosDomar      As Integer

        Dim CanStay          As Boolean

        Dim petType          As Integer

        Dim NroPets          As Integer
    
100     If NpcList(NpcIndex).MaestroUser = UserIndex Then
102         Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

104     With UserList(UserIndex)


106         If .flags.Privilegios And PlayerType.Consejero Then Exit Sub
            
108         If .NroMascotas < MAXMASCOTAS Then

110             If NpcList(NpcIndex).MaestroNPC > 0 Or NpcList(NpcIndex).MaestroUser > 0 Then
112                 Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
                '    Call WriteConsoleMsg(UserIndex, "No puedes domar más de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                '    Exit Sub
                'End If

114             puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))

116             If .clase = eClass.Druid Then
118                 puntosDomar = puntosDomar / 6 'original es 6
                Else
120                 puntosDomar = puntosDomar / 11
                End If

122             If NpcList(NpcIndex).flags.Domable <= puntosDomar And RandomNumber(1, 5) = 1 Then

                    Dim Index As Integer

124                 .NroMascotas = .NroMascotas + 1

126                 Index = FreeMascotaIndex(UserIndex)
128                 .MascotasIndex(Index) = NpcIndex
130                 .MascotasType(Index) = NpcList(NpcIndex).Numero

132                 NpcList(NpcIndex).MaestroUser = UserIndex
                    
                    .flags.ModificoMascotas = True
                    
134                 Call FollowAmo(NpcIndex)
136                 Call ReSpawnNpc(NpcList(NpcIndex))

138                 Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)

                    ' Es zona segura?
140                 If MapInfo(.Pos.Map).Seguro = 1 Then
142                     petType = NpcList(NpcIndex).Numero
144                     NroPets = .NroMascotas

146                     Call QuitarNPC(NpcIndex)

148                     .MascotasType(Index) = petType
150                     .NroMascotas = NroPets

152                     Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

154                 If Not .flags.UltimoMensaje = 5 Then
156                     Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
158                     .flags.UltimoMensaje = 5
                    End If

                End If

160             Call SubirSkill(UserIndex, eSkill.Domar)

            Else
162             Call WriteConsoleMsg(UserIndex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

ErrHandler:
164     Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo PuedeDomarMascota_Err
    
        

        '***************************************************
        'Author: ZaMa
        'This function checks how many NPCs of the same type have
        'been tamed by the user.
        'Returns True if that amount is less than two.
        '***************************************************
        Dim i           As Long

        Dim numMascotas As Long
    
100     For i = 1 To MAXMASCOTAS

102         If UserList(UserIndex).MascotasType(i) = NpcList(NpcIndex).Numero Then
104             numMascotas = numMascotas + 1

            End If

106     Next i
    
108     If numMascotas <= 1 Then PuedeDomarMascota = True
    
        
        Exit Function

PuedeDomarMascota_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeDomarMascota", Erl)

        
End Function
