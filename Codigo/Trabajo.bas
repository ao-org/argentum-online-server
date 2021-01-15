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
116                     .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje

118                     .Char.ShieldAnim = NingunEscudo
120                     .Char.WeaponAnim = NingunArma
122                     .Char.CascoAnim = NingunCasco

124                     Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
126                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

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
134     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoPermanecerOculto", Erl)

136     Resume Next
        
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
                    ByVal slot As Integer)
        
        On Error GoTo DoNavega_Err

100     With UserList(UserIndex)

102         If .Invent.BarcoObjIndex <> .Invent.Object(slot).ObjIndex Then

104             If Not EsGM(UserIndex) Then
            
106                 Select Case Barco.Subtipo
        
                        Case 2  'Galera
        
108                         If .clase <> eClass.Assasin And .clase <> eClass.Pirat And .clase <> eClass.Bandit And .clase <> eClass.Cleric And .clase <> eClass.Thief And .clase <> eClass.Paladin Then
110                             Call WriteConsoleMsg(UserIndex, "Solo los Piratas, Asesinos, Bandidos, Clérigos, Ladrones y Paladines pueden usar Galera!!", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
112                     Case 3  'Galeón
                    
114                         If .clase <> eClass.Thief And .clase <> eClass.Pirat Then
116                             Call WriteConsoleMsg(UserIndex, "Solo los Ladrones y Piratas pueden usar Galeón!!", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
                    End Select
                    
                End If
            
                Dim SkillNecesario As Byte
118             SkillNecesario = IIf(.clase = eClass.Trabajador Or .clase = eClass.Pirat, Barco.MinSkill \ 2, Barco.MinSkill)
            
                ' Tiene el skill necesario?
120             If .Stats.UserSkills(eSkill.Navegacion) < SkillNecesario Then
122                 Call WriteConsoleMsg(UserIndex, "Necesitas al menos " & SkillNecesario & " puntos en navegación para poder usar este " & IIf(Barco.Subtipo = 0, "traje", "barco") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
124             If .Invent.BarcoObjIndex = 0 Then
126                 Call WriteNavigateToggle(UserIndex)
128                 .flags.Navegando = 1
                End If
    
130             .Invent.BarcoObjIndex = .Invent.Object(slot).ObjIndex
132             .Invent.BarcoSlot = slot
    
134             If .flags.Montado > 0 Then
136                 Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
                End If
                
                If .flags.Mimetizado = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Pierdes el efecto del mimetismo.", FontTypeNames.FONTTYPE_INFO)
                    .Counters.Mimetismo = 0
                    .flags.Mimetizado = 0
                End If

138             Call WriteNadarToggle(UserIndex, Barco.Subtipo = 0)
    
140             If .flags.Muerto = 0 Then
142                 .Char.Body = Barco.Ropaje
                
144                 If Barco.Subtipo = 0 Then
146                     .Char.Head = .OrigChar.Head
                    Else
148                     .Char.Head = 0
                    End If
                Else
150                 If Barco.Subtipo = 0 Then
152                     .Char.Body = iRopaBuceoMuerto
154                     .Char.Head = iCabezaMuerto
                    Else
156                     .Char.Body = iFragataFantasmal
158                     .Char.Head = 0
                    End If
                End If
    
160             .Char.ShieldAnim = NingunEscudo
162             .Char.WeaponAnim = NingunArma

164             .Char.speeding = Barco.Velocidad
            
            Else
166             Call WriteNadarToggle(UserIndex, False)
            
168             Call WriteNavigateToggle(UserIndex)
    
170             .flags.Navegando = 0
172             .Invent.BarcoObjIndex = 0
174             .Invent.BarcoSlot = 0
    
176             If .flags.Muerto = 0 Then
178                 .Char.Head = .OrigChar.Head
        
180                 If .Invent.ArmourEqpObjIndex > 0 Then
182                     .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                    Else
184                     Call DarCuerpoDesnudo(UserIndex)
                    End If
        
186                 If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

188                 If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

190                 If .Invent.NudilloObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.NudilloObjIndex).WeaponAnim

192                 If .Invent.HerramientaEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.HerramientaEqpObjIndex).WeaponAnim

194                 If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                
196                 .Char.speeding = VelocidadNormal
                Else
198                 .Char.Body = iCuerpoMuerto
200                 .Char.Head = 0
202                 .Char.ShieldAnim = NingunEscudo
204                 .Char.WeaponAnim = NingunArma
206                 .Char.CascoAnim = NingunCasco
                
208                 .Char.speeding = VelocidadMuerto
                End If

            End If
        
            ' Volver visible
210         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 And .flags.invisible = 0 Then
212             .flags.Oculto = 0
214             .Counters.TiempoOculto = 0

                'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
216             Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
218             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            End If

220         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
222         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
224         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, .Pos.X, .Pos.Y))
    
        End With
        
        Exit Sub

DoNavega_Err:
226     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoNavega", Erl)

228     Resume Next
        
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
        
        On Error GoTo FundirMineral_Err
        

100     If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then

            Dim SkillRequerido As Integer
102         SkillRequerido = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill * ModFundirMineral(UserList(UserIndex).clase)
   
104         If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
                UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) >= SkillRequerido Then
            
106             Call DoLingotes(UserIndex)
        
108         ElseIf SkillRequerido > 100 Then
110             Call WriteConsoleMsg(UserIndex, "Tu clase no tiene el conocimiento suficiente para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
                
            Else
112             Call WriteConsoleMsg(UserIndex, "No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas " & SkillRequerido & " puntos en minería.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

FundirMineral_Err:
114     Call RegistrarError(Err.Number, Err.description, "Trabajo.FundirMineral", Erl)
116     Resume Next
        
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        'Call LogTarea("Sub TieneObjetos")
        
        On Error GoTo TieneObjetos_Err
        

        Dim i     As Long

        Dim Total As Long

100     For i = 1 To UserList(UserIndex).CurrentInventorySlots

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
104             Total = Total + UserList(UserIndex).Invent.Object(i).Amount

            End If

106     Next i

108     If cant <= Total Then
110         TieneObjetos = True
            Exit Function

        End If
        
        
        Exit Function

TieneObjetos_Err:
112     Call RegistrarError(Err.Number, Err.description, "Trabajo.TieneObjetos", Erl)
114     Resume Next
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        'Call LogTarea("Sub QuitarObjetos")
        
        On Error GoTo QuitarObjetos_Err
        
100     With UserList(UserIndex)

            Dim i As Long
    
102         For i = 1 To .CurrentInventorySlots
    
104             If .Invent.Object(i).ObjIndex = ItemIndex Then
    
106                 .Invent.Object(i).Amount = .Invent.Object(i).Amount - cant
    
108                 If .Invent.Object(i).Amount <= 0 Then
110                     If .Invent.Object(i).Equipped Then
112                         Call Desequipar(UserIndex, i)
                        End If
    
114                     cant = Abs(.Invent.Object(i).Amount)
116                     .Invent.Object(i).Amount = 0
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
130     Call RegistrarError(Err.Number, Err.description, "Trabajo.QuitarObjetos", Erl)
132     Resume Next
        
End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
102     If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
104     If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)

        
        Exit Sub

HerreroQuitarMateriales_Err:
106     Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroQuitarMateriales", Erl)
108     Resume Next
        
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)

102     If ObjData(ItemIndex).MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, ObjData(ItemIndex).MaderaElfica, UserIndex)

        
        Exit Sub

CarpinteroQuitarMateriales_Err:
104     Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroQuitarMateriales", Erl)
106     Resume Next
        
End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)

        
        Exit Sub

AlquimistaQuitarMateriales_Err:
102     Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaQuitarMateriales", Erl)
104     Resume Next
        
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreQuitarMateriales_Err
        

100     If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex)
102     If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
104     If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex)

        
        Exit Sub

SastreQuitarMateriales_Err:
106     Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreQuitarMateriales", Erl)
108     Resume Next
        
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
122     Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroTieneMateriales", Erl)
124     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaTieneMateriales", Erl)
114     Resume Next
        
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
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreTieneMateriales", Erl)
134     Resume Next
        
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
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroTieneMateriales", Erl)
134     Resume Next
        
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruir_Err
        
100     PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkHerreria

        
        Exit Function

PuedeConstruir_Err:
102     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruir", Erl)
104     Resume Next
        
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
118     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirHerreria", Erl)
120     Resume Next
        
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub

102     If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
104         Call HerreroQuitarMateriales(UserIndex, ItemIndex)
106         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
108         Call WriteUpdateSta(UserIndex)
            ' AGREGAR FX
    
110         If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el arma!", FontTypeNames.FONTTYPE_INFO)
112             Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5)
114         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el escudo!", FontTypeNames.FONTTYPE_INFO)
116             Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5)
118         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el casco!", FontTypeNames.FONTTYPE_INFO)
120             Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5)
122         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
                'Call WriteConsoleMsg(UserIndex, "Has construido la armadura!", FontTypeNames.FONTTYPE_INFO)
124             Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5)

            End If

            Dim MiObj As obj

126         MiObj.Amount = 1
128         MiObj.ObjIndex = ItemIndex

130         If Not MeterItemEnInventario(UserIndex, MiObj) Then
132             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            'End If
    
134         Call SubirSkill(UserIndex, eSkill.Herreria)
136         Call UpdateUserInv(True, UserIndex, 0)
138         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

140         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

HerreroConstruirItem_Err:
142     Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroConstruirItem", Erl)
144     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirCarpintero", Erl)
112     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirAlquimista", Erl)
112     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirSastre", Erl)
112     Resume Next
        
End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub

102     If CarpinteroTieneMateriales(UserIndex, ItemIndex) _
                And UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria _
                And PuedeConstruirCarpintero(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 5 Then
    
104         If UserList(UserIndex).Stats.MinSta > 2 Then
106             Call QuitarSta(UserIndex, 2)
        
            Else
108             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para trabajar.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If
    
112         Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido un objeto!", FontTypeNames.FONTTYPE_INFO)
            'Call WriteOroOverHead(UserIndex, 1, UserList(UserIndex).Char.CharIndex)
114         Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5)
    
            Dim MiObj As obj

116         MiObj.Amount = 1
118         MiObj.ObjIndex = ItemIndex

120         If Not MeterItemEnInventario(UserIndex, MiObj) Then
122             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            ' End If
    
124         Call SubirSkill(UserIndex, eSkill.Carpinteria)
            'Call UpdateUserInv(True, UserIndex, 0)
126         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

128         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

CarpinteroConstruirItem_Err:
130     Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroConstruirItem", Erl)
132     Resume Next
        
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
    
110         Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
            Dim MiObj As obj

114         MiObj.Amount = 1
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
128     Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaConstruirItem", Erl)
130     Resume Next
        
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
114         Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5)
    '112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, " 1", 5))
            ' UserList(UserIndex).flags.UltimoMensaje = 9
            ' End If
        
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(63, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
            Dim MiObj As obj

118         MiObj.Amount = 1
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
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreConstruirItem", Erl)
134     Resume Next
        
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
116     Call RegistrarError(Err.Number, Err.description, "Trabajo.MineralesParaLingote", Erl)
118     Resume Next
        
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
        '    Call LogTarea("Sub DoLingotes")
        
        On Error GoTo DoLingotes_Err
        

100     If UserList(UserIndex).Stats.MinSta > 5 Then
102         Call QuitarSta(UserIndex, 5)
    
        Else
        
104         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
106         Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub

        End If

        Dim slot As Integer
        Dim obji As Integer

108     slot = UserList(UserIndex).flags.TargetObjInvSlot
110     obji = UserList(UserIndex).Invent.Object(slot).ObjIndex
    
112     Dim cant As Byte: cant = RandomNumber(1, 3)
    
        Dim necesarios As Integer

114     necesarios = MineralesParaLingote(obji, cant)
    
116     If UserList(UserIndex).Invent.Object(slot).Amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
118         Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
120         Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub

        End If
    
122     UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount - MineralesParaLingote(obji, cant)

124     If UserList(UserIndex).Invent.Object(slot).Amount < 1 Then
126         UserList(UserIndex).Invent.Object(slot).Amount = 0
128         UserList(UserIndex).Invent.Object(slot).ObjIndex = 0

        End If
    
        Dim nPos  As WorldPos

        Dim MiObj As obj

130     MiObj.Amount = cant
132     MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex

134     If Not MeterItemEnInventario(UserIndex, MiObj) Then
136         Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

        End If

138     Call UpdateUserInv(False, UserIndex, slot)
        
        
140        Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, cant, 5)
    
        'If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
        '  Call WriteConsoleMsg(UserIndex, "¡Has obtenido lingotes!", FontTypeNames.FONTTYPE_INFO)
            
        '  UserList(UserIndex).flags.UltimoMensaje = 5
        'End If
    
142     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
144     Call SubirSkill(UserIndex, eSkill.Mineria)
  
146     UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
        
148     If UserList(UserIndex).Counters.Trabajando = 1 And Not UserList(UserIndex).flags.UsandoMacro Then
150         Call WriteMacroTrabajoToggle(UserIndex, True)

        End If
    
        
        Exit Sub

DoLingotes_Err:
152     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoLingotes", Erl)
154     Resume Next
        
End Sub

Function ModFundicion(ByVal clase As eClass) As Single
        
        On Error GoTo ModFundicion_Err
        

100     Select Case clase

            Case eClass.Trabajador
102             ModFundicion = 3

104         Case Else
106             ModFundicion = 1

        End Select

        
        Exit Function

ModFundicion_Err:
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModFundicion", Erl)
110     Resume Next
        
End Function

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
112     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModAlquimia", Erl)
114     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModSastre", Erl)
110     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModCarpinteria", Erl)
110     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModHerreriA", Erl)
110     Resume Next
        
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
126     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoAdminInvisible", Erl)

128     Resume Next
        
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

122     If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
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
144         obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
146         Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.Amount & " ramitas.", FontTypeNames.FONTTYPE_INFO)
    
148         Call MakeObj(obj, Map, X, Y)
    
            'Seteamos la fogata como el nuevo TargetObj del user
150         UserList(UserIndex).flags.TargetObj = FOGATA_APAG
        Else

            '[CDT 17-02-2004]
152         If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
154             Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
156             UserList(UserIndex).flags.UltimoMensaje = 10

            End If

            '[/CDT]
        End If

158     Call SubirSkill(UserIndex, Supervivencia)

        
        Exit Sub

TratarDeHacerFogata_Err:
160     Call RegistrarError(Err.Number, Err.description, "Trabajo.TratarDeHacerFogata", Erl)
162     Resume Next
        
End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal RedDePesca As Boolean = False, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte       As Integer
        Dim res          As Integer
        Dim RestaStamina As Byte

100     RestaStamina = IIf(RedDePesca, 2, 1)
    
102     With UserList(UserIndex)
    
104         If .Stats.MinSta > RestaStamina Then
106             Call QuitarSta(UserIndex, RestaStamina)
        
            Else
            
108             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para pescar.", FontTypeNames.FONTTYPE_INFO)
            
110             Call WriteMacroTrabajoToggle(UserIndex, False)
            
                Exit Sub

            End If

            Dim Skill As Integer

112         Skill = .Stats.UserSkills(eSkill.Pescar)
        
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
116         res = RandomNumber(1, Suerte)
    
118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))

120         If res < 6 Then

                Dim nPos  As WorldPos
                Dim MiObj As obj
        
122             MiObj.Amount = IIf(RedDePesca, RandomNumber(2, 5), IIf(ObjetoDorado, RandomNumber(1, 3), 1)) * RecoleccionMult
124             MiObj.ObjIndex = ObtenerPezRandom(2)
        
126             If MiObj.ObjIndex = 0 Then Exit Sub
        
128             If Not MeterItemEnInventario(UserIndex, MiObj) Then
130                 Call TirarItemAlPiso(.Pos, MiObj)
                End If

132             Call WriteRenderValueMsg(UserIndex, .Pos.X, .Pos.Y, MiObj.Amount, 5)
        
                ' Al pescar también podés sacar cosas raras (se setean desde RecursosEspeciales.dat)
                Dim i As Integer

                ' Por cada drop posible
134             For i = 1 To UBound(EspecialesPesca)
                    ' Tiramos al azar entre 1 y la probabilidad
136                 res = RandomNumber(1, IIf(RedDePesca, EspecialesPesca(i).data * 2, EspecialesPesca(i).data)) ' Red de pesca chance x2 (revisar)
            
                    ' Si tiene suerte y le pega
138                 If res = 1 Then
140                     MiObj.ObjIndex = EspecialesPesca(i).ObjIndex
142                     MiObj.Amount = 1 ' Solo un item por vez
                
144                     If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                    
                        ' Le mandamos un mensaje
146                     Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(EspecialesPesca(i).ObjIndex).name & "!", FontTypeNames.FONTTYPE_INFO)

                        ' TODO: Sonido ?
                        'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                    End If

                Next

            End If
    
148         Call SubirSkill(UserIndex, eSkill.Pescar)
    
150         .Counters.Trabajando = .Counters.Trabajando + 1
    
            'Ladder 06/07/14 Activamos el macro de trabajo
152         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
154             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
156     Call LogError("Error en DoPescar. Error " & Err.Number & " - " & Err.description)

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal victimaIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 05/04/2010
        'Last Modification By: ZaMa
        '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
        '27/11/2009: ZaMa - Optimizacion de codigo.
        '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
        '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
        '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
        '23/04/2010: ZaMa - No se puede robar mas sin energia.
        '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
        '*************************************************

        On Error GoTo ErrHandler

        Dim OtroUserIndex As Integer

100     If MapInfo(UserList(victimaIndex).Pos.Map).Seguro = 1 Then Exit Sub
    
102     If UserList(victimaIndex).flags.EnConsulta Then
104         Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
106     With UserList(LadrOnIndex)
    
108         If .flags.Seguro Then
        
110             If Status(LadrOnIndex) = 1 Then
112                 Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            Else

114             If .Faccion.ArmadaReal = 1 Then
            
116                 If Status(victimaIndex) = 1 Then
118                     Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejercito real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub

                    End If

                End If

            End If
        
            ' Caos robando a caos?
120         If UserList(victimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
122             Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legion oscura.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        
124         If TriggerZonaPelea(LadrOnIndex, victimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
            ' Tiene energia?
126         If .Stats.MinSta < 15 Then
        
128             If .genero = eGenero.Hombre Then
130                 Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
                
                Else
132                 Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
                Exit Sub

            End If
        
134         If .GuildIndex > 0 Then
        
136             If .flags.SeguroClan Then
            
138                 If .GuildIndex = UserList(victimaIndex).GuildIndex Then
140                     Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

                End If

            End If

142         If .Grupo.EnGrupo > 0 Then
        
144             If .GuildIndex = UserList(victimaIndex).GuildIndex Then
146                 Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If

148         If .Grupo.EnGrupo = True Then

                Dim i As Byte
150             For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

152                 If UserList(.Grupo.Lider).Grupo.Miembros(i) = victimaIndex Then
154                     Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

156             Next i

            End If
        
            ' Quito energia
158         Call QuitarSta(LadrOnIndex, 15)
                
160         If UserList(victimaIndex).flags.Privilegios And PlayerType.user Then
            
                Dim Suerte     As Integer
                Dim res        As Integer
                Dim RobarSkill As Byte
            
162             RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
164             If RobarSkill <= 10 Then
166                 Suerte = 35
                
168             ElseIf RobarSkill <= 20 Then
170                 Suerte = 30
                
172             ElseIf RobarSkill <= 30 Then
174                 Suerte = 28
                
176             ElseIf RobarSkill <= 40 Then
178                 Suerte = 24
                
180             ElseIf RobarSkill <= 50 Then
182                 Suerte = 22
                
184             ElseIf RobarSkill <= 60 Then
186                 Suerte = 20
                
188             ElseIf RobarSkill <= 70 Then
190                 Suerte = 18
                
192             ElseIf RobarSkill <= 80 Then
194                 Suerte = 15
                
196             ElseIf RobarSkill <= 90 Then
198                 Suerte = 10
                
200             ElseIf RobarSkill < 100 Then
202                 Suerte = 7
                
                Else
204                 Suerte = 5
                
                End If
            
206             res = RandomNumber(1, Suerte)
                
208             If res < 3 Then 'Exito robo
                
210                 If UserList(victimaIndex).flags.Comerciando Then
212                     OtroUserIndex = UserList(victimaIndex).ComUsu.DestUsu
                        
214                     If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
216                         Call WriteConsoleMsg(victimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_TALK)
218                         Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
220                         Call LimpiarComercioSeguro(victimaIndex)

                        End If

                    End If
               
222                 If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    
224                     If TieneObjetosRobables(victimaIndex) Then
226                         Call RobarObjeto(LadrOnIndex, victimaIndex)
                        Else
228                         Call WriteConsoleMsg(LadrOnIndex, UserList(victimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else 'Roba oro

230                     If UserList(victimaIndex).Stats.GLD > 0 Then

                            Dim n As Long

232                         If .clase = eClass.Thief Then

                                ' Si no tine puestos los guantes de hurto roba un 50% menos.
234                             If .Invent.DañoMagicoEqpObjIndex > 0 Then
236                                 If ObjData(.Invent.DañoMagicoEqpObjIndex).Subtipo = 5 Then
238                                     n = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100) * OroMult
                                    Else
240                                     n = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50) * OroMult
                                    End If
                                Else
242                                 n = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50) * OroMult
                                End If
    
                            Else
244                             n = RandomNumber(1, 100)
    
                            End If

246                         If n > UserList(victimaIndex).Stats.GLD Then n = UserList(victimaIndex).Stats.GLD
                        
248                         UserList(victimaIndex).Stats.GLD = UserList(victimaIndex).Stats.GLD - n
                        
250                         .Stats.GLD = .Stats.GLD + n

252                         If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO
                        
254                         Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & PonerPuntos(n) & " monedas de oro a " & UserList(victimaIndex).name, FontTypeNames.FONTTYPE_INFO)
256                         Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
258                         Call WriteUpdateGold(victimaIndex) 'Le actualizamos la billetera a la victima
                        Else
260                         Call WriteConsoleMsg(LadrOnIndex, UserList(victimaIndex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                
262                 Call SubirSkill(LadrOnIndex, eSkill.Robar)
            
                Else
264                 Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
266                 Call WriteConsoleMsg(victimaIndex, "" & .name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
268                 Call SubirSkill(LadrOnIndex, eSkill.Robar)

                End If
            
270             If Status(LadrOnIndex) = 1 Then Call VolverCriminal(LadrOnIndex)
        
272             If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)
            
                'If Not Criminal(LadrOnIndex) Then
                    'If Not Criminal(VictimaIndex) Then
                        'Call VolverCriminal(LadrOnIndex)
                    'End If
                'End If
            
                ' Se pudo haber convertido si robo a un ciuda
                'If Criminal(LadrOnIndex) Then
                    '.Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron
                    'If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP
                'End If

            End If

        End With

        Exit Sub

ErrHandler:
274     Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

Public Function ObjEsRobable(ByVal victimaIndex As Integer, ByVal slot As Integer) As Boolean
        ' Agregué los barcos
        ' Esta funcion determina qué objetos son robables.
        
        On Error GoTo ObjEsRobable_Err
        

        Dim OI As Integer

100     OI = UserList(victimaIndex).Invent.Object(slot).ObjIndex

102     ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(victimaIndex).Invent.Object(slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).donador = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otRunas And ObjData(OI).Instransferible = 0 And ObjData(OI).OBJType <> eOBJType.otMonturas

        
        Exit Function

ObjEsRobable_Err:
104     Call RegistrarError(Err.Number, Err.description, "Trabajo.ObjEsRobable", Erl)
106     Resume Next
        
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal victimaIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
        '***************************************************
        
        On Error GoTo RobarObjeto_Err
    
        

        Dim flag As Boolean
        Dim i    As Integer

100     flag = False

102     With UserList(victimaIndex)

104         If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
106             i = 1

108             Do While Not flag And i <= .CurrentInventorySlots

                    'Hay objeto en este slot?
110                 If .Invent.Object(i).ObjIndex > 0 Then
                
112                     If ObjEsRobable(victimaIndex, i) Then
                    
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
                
124                     If ObjEsRobable(victimaIndex, i) Then
                    
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
        
132             ObjAmount = .Invent.Object(i).Amount
        
                'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
134             num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
136             MiObj.Amount = num
138             MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
140             .Invent.Object(i).Amount = ObjAmount - num
                    
142             If .Invent.Object(i).Amount <= 0 Then
144                 Call QuitarUserInvItem(victimaIndex, CByte(i), 1)

                End If
                
146             Call UpdateUserInv(False, victimaIndex, CByte(i))
                    
148             If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
150                 Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
                
                End If
        
152             If UserList(LadrOnIndex).clase = eClass.Thief Then
154                 Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
                
                Else
156                 Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
                
                End If

            Else
158             Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningun objeto.", FontTypeNames.FONTTYPE_INFO)

            End If

            'If exiting, cancel de quien es robado
160         Call CancelExit(victimaIndex)

        End With

        
        Exit Sub

RobarObjeto_Err:
162     Call RegistrarError(Err.Number, Err.description, "Trabajo.RobarObjeto", Erl)

        
End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
        
        On Error GoTo DoApuñalar_Err
        

        '***************************************************
        'Autor: Nacho (Integer) & Unknown (orginal version)
        'Last Modification: 04/17/08 - (NicoNZ)
        'Simplifique la cuenta que hacia para sacar la suerte
        'y arregle la cuenta que hacia para sacar el daño
        '***************************************************
        Dim Suerte As Integer

        Dim Skill  As Integer
    
100     Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)
    
102     Select Case UserList(UserIndex).clase

            Case eClass.Assasin '25
104             Suerte = 0.2 * Skill + 5

106         Case eClass.Pirat, eClass.Hunter '20
108             Suerte = 0.155 * Skill + 4.5

110         Case Else ' 15
112             Suerte = 0.11 * Skill + 4

        End Select
        
        ' Daga especial da +5 de prob. de apu
114     If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Subtipo = 42 Then
116         Suerte = Suerte + 5
        End If
    
118     If RandomNumber(1, 100) < Suerte Then
120         If VictimUserIndex <> 0 Then
122             daño = daño * 1.5
            
124             UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño

126             Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageEfectOverHead("¡" & daño & "!", UserList(VictimUserIndex).Char.CharIndex, vbYellow))

128             If UserList(UserIndex).ChatCombate = 1 Then
                    'Call WriteEfectOverHead(UserIndex, daño, UserList(UserIndex).Char.CharIndex) 'LADDER 21.11.08
130                 Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

                End If

132             If UserList(VictimUserIndex).ChatCombate = 1 Then
134                 Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
            
            Else
136             daño = daño * 2

138             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño

140             If UserList(UserIndex).ChatCombate = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 1.5), FontTypeNames.FONTTYPE_FIGHT)
142                 Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, daño)

                End If
            
144             Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageEfectOverHead("¡" & daño & "!", Npclist(VictimNpcIndex).Char.CharIndex, vbYellow))

                '[Alejo]
146             Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)

            End If

        Else

148         If UserList(UserIndex).ChatCombate = 1 Then
150             Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If
    
152     Call SubirSkill(UserIndex, Apuñalar)

        
        Exit Sub

DoApuñalar_Err:
154     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoApuñalar", Erl)
156     Resume Next
        
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
        
        On Error GoTo DoGolpeCritico_Err

        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 28/01/2007
        '***************************************************
        Dim Suerte As Integer
        Dim Skill  As Integer
        
100     With UserList(UserIndex)
        
102         If .clase <> eClass.Bandit Then Exit Sub
104         If .Invent.WeaponEqpSlot = 0 Then Exit Sub
106         If ObjData(.Invent.WeaponEqpObjIndex).Subtipo <> 2 Then Exit Sub

108         Skill = .Stats.UserSkills(eSkill.Wrestling)
        
        End With

110     Suerte = 0.2 * Skill + 5

112     If VictimNpcIndex <> 0 Then Suerte = Suerte * 2

114     If RandomNumber(1, 100) < Suerte Then
116         daño = Int(daño * 0.5)

118         If VictimUserIndex <> 0 Then
120             UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
122             Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
124             Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
                
            Else
126             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
128             Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
130             Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)

            End If

        End If
        
        Exit Sub

DoGolpeCritico_Err:
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoGolpeCritico", Erl)

134     Resume Next
        
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarSta_Err
        
100     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

102     If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
104     If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub
106     Call WriteUpdateSta(UserIndex)

        
        Exit Sub

QuitarSta_Err:
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.QuitarSta", Erl)
110     Resume Next
        
End Sub

Public Sub DoRaices(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
    
100     With UserList(UserIndex)
    
102         If .Stats.MinSta > 2 Then
104             Call QuitarSta(UserIndex, 2)
        
            Else
            
106             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para obtener raices.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            Dim Skill As Integer
110             Skill = .Stats.UserSkills(eSkill.Alquimia)
        
112         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
114         res = RandomNumber(1, Suerte)
    
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
    
            Rem Ladder 06/08/14 Subo un poco la probabilidad de sacar raices... porque era muy lento
118         If res < 7 Then
    
                Dim nPos  As WorldPos
                Dim MiObj As obj
        
                'If .clase = eClass.Druid Then
                'MiObj.Amount = RandomNumber(6, 8)
                ' Else
120             MiObj.Amount = RandomNumber(5, 7)
                ' End If
       
122             If ObjData(.Invent.HerramientaEqpObjIndex).donador = 1 Then
124                 MiObj.Amount = MiObj.Amount * 2
                End If
       
126             MiObj.Amount = MiObj.Amount * RecoleccionMult
128             MiObj.ObjIndex = Raices
        
130             MapData(.Pos.Map, X, Y).ObjInfo.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount - MiObj.Amount
    
132             If MapData(.Pos.Map, X, Y).ObjInfo.Amount < 0 Then
134                 MapData(.Pos.Map, X, Y).ObjInfo.Amount = 0
    
136                 Call AgregarItemLimpieza(.Pos.Map, X, Y)
                
                End If
        
138             If Not MeterItemEnInventario(UserIndex, MiObj) Then
            
140                 Call TirarItemAlPiso(.Pos, MiObj)
            
                End If
        
                'Call WriteConsoleMsg(UserIndex, "¡Has conseguido algunas raices!", FontTypeNames.FONTTYPE_INFO)
142             Call WriteRenderValueMsg(UserIndex, .Pos.X, .Pos.Y, MiObj.Amount, 5)
144             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(60, .Pos.X, .Pos.Y))
            Else
146             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(61, .Pos.X, .Pos.Y))
    
                '[CDT 17-02-2004]
148             If Not .flags.UltimoMensaje = 8 Then
150                 Call WriteConsoleMsg(UserIndex, "¡No has obtenido raices!", FontTypeNames.FONTTYPE_INFO)
152                 .flags.UltimoMensaje = 8
    
                End If
        
                '[/CDT]
            End If
    
154         Call SubirSkill(UserIndex, eSkill.Alquimia)
    
156         .Counters.Trabajando = .Counters.Trabajando + 1
    
158         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
160             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
162     Call LogError("Error en DoRaices")

End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
    
100     With UserList(UserIndex)
    
                'EsfuerzoTalarLeñador = 1
102         If .Stats.MinSta > 2 Then
104             Call QuitarSta(UserIndex, 2)
        
            Else
106             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para talar.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            Dim Skill As Integer
    
110         Skill = .Stats.UserSkills(eSkill.Talar)
112         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
114         res = RandomNumber(1, Suerte)
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        
118         If res < 6 Then
    
                Dim nPos  As WorldPos
    
                Dim MiObj As obj
            
120             Call ActualizarRecurso(.Pos.Map, X, Y)
122             MapData(.Pos.Map, X, Y).ObjInfo.data = GetTickCount() ' Ultimo uso
    
124             MiObj.Amount = IIf(ObjetoDorado, RandomNumber(1, 5), 1) * RecoleccionMult

126             If ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex).Elfico = 0 Then
128                 MiObj.ObjIndex = Leña
                Else
130                 MiObj.ObjIndex = LeñaElfica
                End If

            
132             If MiObj.Amount > MapData(.Pos.Map, X, Y).ObjInfo.Amount Then
134                 MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                End If
            
136             MapData(.Pos.Map, X, Y).ObjInfo.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount - MiObj.Amount
            
138             If Not MeterItemEnInventario(UserIndex, MiObj) Then
140                 Call TirarItemAlPiso(.Pos, MiObj)
                End If
    
                'If Not .flags.UltimoMensaje = 5 Then
                ' Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
                '        .flags.UltimoMensaje = 5
                ' End If
142             Call WriteRenderValueMsg(UserIndex, .Pos.X, .Pos.Y, MiObj.Amount, 5)
144             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
            
                ' Al talar también podés dropear cosas raras (se setean desde RecursosEspeciales.dat)
                Dim i As Integer
    
                ' Por cada drop posible
146             For i = 1 To UBound(EspecialesTala)
                    ' Tiramos al azar entre 1 y la probabilidad
148                 res = RandomNumber(1, EspecialesTala(i).data)
                
                    ' Si tiene suerte y le pega
150                 If res = 1 Then
152                     MiObj.ObjIndex = EspecialesTala(i).ObjIndex
154                     MiObj.Amount = 1 ' Solo un item por vez
                    
                        'If Not MeterItemEnInventario(Userindex, MiObj) Then _
                        'Call TirarItemAlPiso(.Pos, MiObj)
    
                        ' Tiro siempre el item al piso, me parece más rolero, como que cae del árbol :P
156                     Call TirarItemAlPiso(.Pos, MiObj)
    
                        ' Oculto el mensaje porque el item cae al piso
                        'Call WriteConsoleMsg(Userindex, "¡Has conseguido " & ObjData(EspecialesTala(i).ObjIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
                        ' TODO: Sonido ?
                        'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                    End If
    
                Next
        
            Else
                '[CDT 17-02-2004]
158             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(64, .Pos.X, .Pos.Y))
    
160             If Not .flags.UltimoMensaje = 8 Then
162                 Call WriteConsoleMsg(UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
164                 .flags.UltimoMensaje = 8
    
                End If
    
                '[/CDT]
            End If
        
166         Call SubirSkill(UserIndex, eSkill.Talar)
        
168         .Counters.Trabajando = .Counters.Trabajando + 1
    
170         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
172             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With

        Exit Sub

ErrHandler:
174     Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte      As Integer
        Dim res         As Integer
        Dim Metal       As Integer
        Dim Yacimiento  As ObjData

100     With UserList(UserIndex)
    
            'Por Ladder 06/07/2014 Cuando la estamina llega a 0 , el macro se desactiva
102         If .Stats.MinSta > 2 Then
104             Call QuitarSta(UserIndex, 2)
            Else
106             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            'Por Ladder 06/07/2014
    
            Dim Skill As Integer
    
110         Skill = .Stats.UserSkills(eSkill.Mineria)
112         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
114         res = RandomNumber(1, Suerte)
        
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        
118         If res <= 5 Then
    
                Dim MiObj As obj
                Dim nPos  As WorldPos
            
120             Call ActualizarRecurso(.Pos.Map, X, Y)
122             MapData(.Pos.Map, X, Y).ObjInfo.data = GetTickCount() ' Ultimo uso

124             Yacimiento = ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex)
            
126             MiObj.ObjIndex = Yacimiento.MineralIndex
128             MiObj.Amount = IIf(ObjetoDorado, RandomNumber(1, 6), 1) * RecoleccionMult
            
130             If MiObj.Amount > MapData(.Pos.Map, X, Y).ObjInfo.Amount Then
132                 MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                End If
            
134             MapData(.Pos.Map, X, Y).ObjInfo.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount - MiObj.Amount
        
136             If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
            
138             Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
140             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(15, .Pos.X, .Pos.Y))
            
                ' Al minar también puede dropear una gema
                Dim i As Integer
    
                ' Por cada drop posible
142             For i = 1 To Yacimiento.CantItem
                    ' Tiramos al azar entre 1 y la probabilidad
144                 res = RandomNumber(1, Yacimiento.Item(i).Amount)
                
                    ' Si tiene suerte y le pega
146                 If res = 1 Then
                        ' Se lo metemos al inventario (o lo tiramos al piso)
148                     MiObj.ObjIndex = Yacimiento.Item(i).ObjIndex
150                     MiObj.Amount = 1 ' Solo una gema por vez
                    
152                     If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                        
                        ' Le mandamos un mensaje
154                     Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(Yacimiento.Item(i).ObjIndex).name & "!", FontTypeNames.FONTTYPE_INFO)
                        ' TODO: Sonido de drop de gema :P
                        'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                        
                        ' Como máximo dropea una gema
                        'Exit For ' Lo saco a pedido de Haracin
                    End If
    
                Next
            
            Else
156             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(62, .Pos.X, .Pos.Y))
    
                '[CDT 17-02-2004]
158             If Not .flags.UltimoMensaje = 9 Then
                
160                 Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                
162                 .flags.UltimoMensaje = 9
    
                End If
    
                '[/CDT]
            End If
        
164         Call SubirSkill(UserIndex, eSkill.Mineria)
        
166         .Counters.Trabajando = .Counters.Trabajando + 1
        
168         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
170             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
172     Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
        
        On Error GoTo DoMeditar_Err

        Dim Mana As Long

100     With UserList(UserIndex)

102         .Counters.TimerMeditar = .Counters.TimerMeditar + 1

104         If .Counters.TimerMeditar >= IntervaloMeditar Then

106             Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, 50 + .Stats.UserSkills(eSkill.Meditar) * 0.5))

108             If Mana <= 0 Then Mana = 1

110             .Stats.MinMAN = .Stats.MinMAN + Mana

112             If .Stats.MinMAN >= .Stats.MaxMAN Then
114                 .Stats.MinMAN = .Stats.MaxMAN
116                 .flags.Meditando = False
118                 .Char.FX = 0
                    
120                 Call WriteUpdateMana(UserIndex)
122                 Call SubirSkill(UserIndex, Meditar)

124                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
                
                Else
126                 Call WriteUpdateMana(UserIndex)
128             Call SubirSkill(UserIndex, Meditar)
                End If

130             .Counters.TimerMeditar = 0
            End If

        End With

        
        Exit Sub

DoMeditar_Err:
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoMeditar", Erl)
134     Resume Next
        
End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modif: 15/04/2010
        'Unequips either shield, weapon or helmet from target user.
        '***************************************************
        
        On Error GoTo DoDesequipar_Err
    
        

        Dim Probabilidad   As Integer
        Dim Resultado      As Integer
        Dim WrestlingSkill As Byte
        Dim AlgoEquipado   As Boolean
    
100     With UserList(UserIndex)

            ' Si no tiene guantes de hurto no desequipa.
            'If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
            ' Si no esta solo con manos, no desequipa tampoco.
102         If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
104         WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
106         Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66

        End With
   
108     With UserList(VictimIndex)

            ' Si tiene escudo, intenta desequiparlo
110         If .Invent.EscudoEqpObjIndex > 0 Then
            
112             Resultado = RandomNumber(1, 100)
            
114             If Resultado <= Probabilidad Then
                    ' Se lo desequipo
116                 Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
118                 Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
120                 If .Stats.ELV < 20 Then
122                     Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)

                    End If
                
                    Exit Sub

                End If
            
124             AlgoEquipado = True

            End If
        
            ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
126         If .Invent.WeaponEqpObjIndex > 0 Then
            
128             Resultado = RandomNumber(1, 100)
            
130             If Resultado <= Probabilidad Then
                    ' Se lo desequipo
132                 Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
134                 Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
136                 If .Stats.ELV < 20 Then
138                     Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

                    End If
                
                    Exit Sub

                End If
            
140             AlgoEquipado = True

            End If
        
            ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
142         If .Invent.CascoEqpObjIndex > 0 Then
            
144             Resultado = RandomNumber(1, 100)
            
146             If Resultado <= Probabilidad Then
                    ' Se lo desequipo
148                 Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
150                 Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
152                 If .Stats.ELV < 20 Then
154                     Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)

                    End If
                
                    Exit Sub

                End If
            
156             AlgoEquipado = True

            End If
    
158         If AlgoEquipado Then
160             Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
            Else
162             Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningun item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
        End With

        
        Exit Sub

DoDesequipar_Err:
164     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoDesequipar", Erl)

        
End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal victimaIndex As Integer)
        
        On Error GoTo DoHurtar_Err
    
        

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modif: 03/03/2010
        'Implements the pick pocket skill of the Bandit :)
        '03/03/2010 - Pato: Solo se puede hurtar si no esta en trigger 6 :)
        '***************************************************
        Dim OtroUserIndex As Integer

100     If TriggerZonaPelea(UserIndex, victimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

102     If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub

        Dim res As Integer

104     res = RandomNumber(1, 100)

106     If (res < 20) Then
108         If TieneObjetosRobables(victimaIndex) Then
    
110             If UserList(victimaIndex).flags.Comerciando Then
112                 OtroUserIndex = UserList(victimaIndex).ComUsu.DestUsu
                
114                 If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
116                     Call WriteConsoleMsg(victimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_WARNING)
118                     Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_WARNING)
                
120                     Call LimpiarComercioSeguro(victimaIndex)

                    End If

                End If
                
122             Call RobarObjeto(UserIndex, victimaIndex)
124             Call WriteConsoleMsg(victimaIndex, "" & UserList(UserIndex).name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
            Else
126             Call WriteConsoleMsg(UserIndex, UserList(victimaIndex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

DoHurtar_Err:
128     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoHurtar", Erl)

        
End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal victimaIndex As Integer)
        
        On Error GoTo DoHandInmo_Err
    
        

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modif: 17/02/2007
        'Implements the special Skill of the Thief
        '***************************************************
100     If UserList(victimaIndex).flags.Paralizado = 1 Then Exit Sub
        
        Dim res As Integer

102     res = RandomNumber(0, 100)

104     If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
106         UserList(victimaIndex).flags.Paralizado = 1
108         UserList(victimaIndex).Counters.Paralisis = IntervaloParalizado / 4
        
            'UserList(VictimaIndex).flags.ParalizedByIndex = Userindex
            'UserList(VictimaIndex).flags.ParalizedBy = UserList(Userindex).name
        
110         Call WriteParalizeOK(victimaIndex)
112         Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmovil a tu oponente", FontTypeNames.FONTTYPE_FIGHT)
114         Call WriteConsoleMsg(victimaIndex, "¡El golpe te ha dejado inmovil!", FontTypeNames.FONTTYPE_FIGHT)

        End If

        
        Exit Sub

DoHandInmo_Err:
116     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoHandInmo", Erl)

        
End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
        
        On Error GoTo Desarmar_Err
        

        Dim Suerte As Integer

        Dim res    As Integer

100     If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
102         Suerte = 35
104     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
106         Suerte = 30
108     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
110         Suerte = 28
112     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
114         Suerte = 24
116     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
118         Suerte = 22
120     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
122         Suerte = 20
124     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
126         Suerte = 18
128     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
130         Suerte = 15
132     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
134         Suerte = 10
136     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
138         Suerte = 7
140     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
142         Suerte = 5

        End If

144     res = RandomNumber(1, Suerte)

146     If res <= 2 Then
148         Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
150         Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

152         If UserList(VictimIndex).Stats.ELV < 20 Then
154             Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        

        End If

        
        Exit Sub

Desarmar_Err:
156     Call RegistrarError(Err.Number, Err.description, "Trabajo.Desarmar", Erl)
158     Resume Next
        
End Sub

Public Sub DoMontar(ByVal UserIndex As Integer, ByRef Montura As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoMontar_Err
        

100     If Not CheckRazaTipo(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then
102         Call WriteConsoleMsg(UserIndex, "Tu raza no te permite usar esta montura.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If Not CheckClaseTipo(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then
106         Call WriteConsoleMsg(UserIndex, "Tu clase no te permite usar esta montura.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

108     If UserList(UserIndex).Stats.UserSkills(eSkill.equitacion) < Montura.MinSkill Then
            'Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar esta montura.", FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Ladder 21/11/08
112     If UserList(UserIndex).flags.Montado = 0 Then
114         If (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger > 10) Then
116             Call WriteConsoleMsg(UserIndex, "No podés montar aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

118     If UserList(UserIndex).flags.Meditando Then
120         UserList(UserIndex).flags.Meditando = False
122         UserList(UserIndex).Char.FX = 0
124         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
        End If

126     If UserList(UserIndex).flags.Montado = 1 Then
128         If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
130             If ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
132                 Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MonturaSlot)
                End If

            End If

        End If

134     UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
136     UserList(UserIndex).Invent.MonturaSlot = slot

138     If UserList(UserIndex).flags.Montado = 0 Then
            
140         UserList(UserIndex).Char.Body = Montura.Ropaje

            'UserList(UserIndex).Char.body = Montura.Ropaje
142         UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
144         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
146         UserList(UserIndex).Char.WeaponAnim = NingunArma
148         UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).Char.CascoAnim
150         UserList(UserIndex).flags.Montado = 1
152         UserList(UserIndex).Char.speeding = VelocidadMontura
        Else
154         UserList(UserIndex).flags.Montado = 0
156         UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
158         UserList(UserIndex).Char.speeding = VelocidadNormal

160         If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
162             UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

            Else
164             Call DarCuerpoDesnudo(UserIndex)

            End If
            
166         If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

168         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

170         If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

        End If

172     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

174     Call UpdateUserInv(False, UserIndex, slot)
176     Call WriteEquiteToggle(UserIndex)
178     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

        
        Exit Sub

DoMontar_Err:
180     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoMontar", Erl)
182     Resume Next
        
End Sub

Public Function ApuñalarFunction(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer) As Integer
        
        On Error GoTo ApuñalarFunction_Err
        

        '***************************************************
        'Autor: Nacho (Integer) & Unknown (orginal version)
        'Last Modification: 04/17/08 - (NicoNZ)
        'Simplifique la cuenta que hacia para sacar la suerte
        'y arregle la cuenta que hacia para sacar el daño
        '***************************************************
        Dim Suerte As Integer

        Dim Skill  As Integer

        Dim Random As Byte

100     Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

102     Select Case UserList(UserIndex).clase

            Case eClass.Assasin '35
104             Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 5)
        
106             If VictimNpcIndex = 0 Then
108                 If UserList(VictimUserIndex).Char.Heading = UserList(UserIndex).Char.Heading Then
110                     Random = RandomNumber(1, 3)

112                     If Random = 1 Then
114                         Suerte = 70

                        End If

                    End If

                End If
    
116         Case eClass.Cleric, eClass.Paladin ' 15
118             Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
120         Case eClass.Bard, eClass.Druid '13
122             Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
124         Case Else '8
126             Suerte = Int(0.0361 * Skill + 4.39)

        End Select

128     If RandomNumber(0, 99) < Suerte Then
130         If VictimUserIndex <> 0 Then
132                 ApuñalarFunction = daño * ModicadorApuñalarClase(UserList(UserIndex).clase)
            Else ' A los npcs el apuñalar es un 20% mas
134                ApuñalarFunction = (daño * ModicadorApuñalarClase(UserList(UserIndex).clase)) * 1.2
136                Debug.Print ApuñalarFunction
            End If
        End If
        
        Exit Function

ApuñalarFunction_Err:
138     Call RegistrarError(Err.Number, Err.description, "Trabajo.ApuñalarFunction", Erl)
140     Resume Next
        
End Function

Public Sub ActualizarRecurso(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo ActualizarRecurso_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex

        Dim TiempoActual As Long

102     TiempoActual = GetTickCount()

        ' Data = Ultimo uso
104     If (TiempoActual - MapData(Map, X, Y).ObjInfo.data) * 0.001 > ObjData(ObjIndex).TiempoRegenerar Then
106         MapData(Map, X, Y).ObjInfo.Amount = ObjData(ObjIndex).VidaUtil
108         MapData(Map, X, Y).ObjInfo.data = &H7FFFFFFF   ' Ultimo uso = Max Long

        End If

        
        Exit Sub

ActualizarRecurso_Err:
110     Call RegistrarError(Err.Number, Err.description, "Trabajo.ActualizarRecurso", Erl)
112     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.ObtenerPezRandom", Erl)
110     Resume Next
        
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
116     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModDomar", Erl)

        
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
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.FreeMascotaIndex", Erl)

        
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
    
100     If Npclist(NpcIndex).MaestroUser = UserIndex Then
102         Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

104     With UserList(UserIndex)

106         If .NroMascotas < MAXMASCOTAS Then

108             If Npclist(NpcIndex).MaestroNPC > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
110                 Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
                '    Call WriteConsoleMsg(UserIndex, "No puedes domar más de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                '    Exit Sub
                'End If

112             puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))

114             If .clase = eClass.Druid Then
                    puntosDomar = puntosDomar / 6
                Else
122                 puntosDomar = puntosDomar / 11
                End If

124             If Npclist(NpcIndex).flags.Domable <= puntosDomar And RandomNumber(1, 5) = 1 Then

                    Dim index As Integer

126                 .NroMascotas = .NroMascotas + 1
128                 index = FreeMascotaIndex(UserIndex)
130                 .MascotasIndex(index) = NpcIndex
132                 .MascotasType(index) = Npclist(NpcIndex).Numero

134                 Npclist(NpcIndex).MaestroUser = UserIndex

136                 Call FollowAmo(NpcIndex)
138                 Call ReSpawnNpc(Npclist(NpcIndex))

140                 Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)

                    ' Es zona segura?
142                 If MapInfo(.Pos.Map).Seguro = 1 Then
144                     petType = Npclist(NpcIndex).Numero
146                     NroPets = .NroMascotas

148                     Call QuitarNPC(NpcIndex)

150                     .MascotasType(index) = petType
152                     .NroMascotas = NroPets

154                     Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

156                 If Not .flags.UltimoMensaje = 5 Then
158                     Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
160                     .flags.UltimoMensaje = 5
                    End If

                End If

162             Call SubirSkill(UserIndex, eSkill.Domar)

            Else
164             Call WriteConsoleMsg(UserIndex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

ErrHandler:
166     Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

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

102         If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
104             numMascotas = numMascotas + 1

            End If

106     Next i
    
108     If numMascotas <= 1 Then PuedeDomarMascota = True
    
        
        Exit Function

PuedeDomarMascota_Err:
110     Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeDomarMascota", Erl)

        
End Function

Private Function ModFundirMineral(ByVal clase As eClass) As Integer
        
        On Error GoTo ModFundirMineral_Err
    
        
    
100     If clase = eClass.Trabajador Then
102         ModFundirMineral = 1
        Else
104         ModFundirMineral = 3
        End If
    
        
        Exit Function

ModFundirMineral_Err:
106     Call RegistrarError(Err.Number, Err.description, "Trabajo.ModFundirMineral", Erl)

        
End Function
