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

Public Sub DoPermanecerOculto(ByVal Userindex As Integer)
        '********************************************************
        'Autor: Nacho (Integer)
        'Last Modif: 28/01/2007
        'Chequea si ya debe mostrarse
        'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
        '********************************************************
        
        On Error GoTo DoPermanecerOculto_Err
    
100     With UserList(Userindex)
    
102         .Counters.TiempoOculto = .Counters.TiempoOculto - 1

104         If .Counters.TiempoOculto <= 0 Then

106             .Counters.TiempoOculto = 0
108             .flags.Oculto = 0

110             If .flags.Navegando = 1 Then
            
112                 If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
114                     .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje

116                     .Char.ShieldAnim = NingunEscudo
118                     .Char.WeaponAnim = NingunArma
120                     .Char.CascoAnim = NingunCasco

122                     Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
124                     Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

                    End If

                Else

126                 If .flags.invisible = 0 Then
128                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
130                     Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
            End If
    
        End With

        Exit Sub

DoPermanecerOculto_Err:
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoPermanecerOculto", Erl)

134     Resume Next
        
End Sub

Public Sub DoOcultarse(ByVal Userindex As Integer)

        'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
        'Modifique la fórmula y ahora anda bien.
        On Error GoTo ErrHandler

        Dim Suerte As Double
        Dim res    As Integer
        Dim Skill  As Integer
    
100     With UserList(Userindex)
    
102         Skill = .Stats.UserSkills(eSkill.Ocultarse)
104         Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
106         res = RandomNumber(1, 100)

108         If res <= Suerte Then

110             .flags.Oculto = 1
112             Suerte = (-0.000001 * (100 - Skill) ^ 3)
114             Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
116             Suerte = Suerte + (-0.0088 * (100 - Skill))
118             Suerte = Suerte + (0.9571)
120             Suerte = Suerte * IntervaloOculto
        
122             If .clase = eClass.Bandit Then
124                 .Counters.TiempoOculto = Int(Suerte / 2)
                Else
126                 .Counters.TiempoOculto = Suerte

                End If
    
128             If .flags.AnilloOcultismo = 1 Then
130                 .Counters.TiempoOculto = Suerte * 3
                Else
132                 .Counters.TiempoOculto = Suerte

                End If
  
134             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, True))

                'Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
136             Call WriteLocaleMsg(Userindex, "55", FontTypeNames.FONTTYPE_INFO)
138             Call SubirSkill(Userindex, Ocultarse)
            Else

140             If Not .flags.UltimoMensaje = 4 Then
                    'Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
142                 Call WriteLocaleMsg(Userindex, "57", FontTypeNames.FONTTYPE_INFO)
144                 .flags.UltimoMensaje = 4

                End If

            End If

146         .Counters.Ocultando = .Counters.Ocultando + 1
    
        End With

        Exit Sub

ErrHandler:
148     Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNadar(ByVal Userindex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoNadar_Err
        

        Dim ModNave As Long

100     If UserList(Userindex).flags.Nadando = 0 Then
    
102         If UserList(Userindex).flags.Muerto = 0 Then
                '(Nacho)
    
104             UserList(Userindex).Char.Body = 694
                'If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGalera
                'If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleon
            Else
106             UserList(Userindex).Char.Body = iFragataFantasmal

            End If
    
108         UserList(Userindex).Char.ShieldAnim = NingunEscudo
110         UserList(Userindex).Char.WeaponAnim = NingunArma
112         UserList(Userindex).Char.CascoAnim = NingunCasco
114         UserList(Userindex).flags.Nadando = 1
    
        Else
    
116         UserList(Userindex).flags.Nadando = 0
    
118         If UserList(Userindex).flags.Muerto = 0 Then
120             UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
        
122             If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
124                 UserList(Userindex).Char.Body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje
            
                Else
126                 Call DarCuerpoDesnudo(Userindex)

                End If
        
128             If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then UserList(Userindex).Char.ShieldAnim = ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).ShieldAnim

130             If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then UserList(Userindex).Char.WeaponAnim = ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).WeaponAnim

132             If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then UserList(Userindex).Char.CascoAnim = ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).CascoAnim
            Else
134             UserList(Userindex).Char.Body = iCuerpoMuerto
136             UserList(Userindex).Char.Head = iCabezaMuerto
138             UserList(Userindex).Char.ShieldAnim = NingunEscudo
140             UserList(Userindex).Char.WeaponAnim = NingunArma
142             UserList(Userindex).Char.CascoAnim = NingunCasco

            End If

        End If

144     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        'Call WriteNadarToggle(UserIndex)
146     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

        
        Exit Sub

DoNadar_Err:
148     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoNadar", Erl)
150     Resume Next
        
End Sub

Public Sub DoNavega(ByVal Userindex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal slot As Integer)
        
        On Error GoTo DoNavega_Err

100     With UserList(Userindex)
    
            ' Acordate que el Trabajador solo necesita 60 de Navegacion para usar barca!
            Dim SkillNecesario As Byte
102             SkillNecesario = IIf(.clase = eClass.Trabajador, 35, Barco.MinSkill)
        
            ' Tiene el skill necesario?
104         If .Stats.UserSkills(eSkill.Navegacion) < SkillNecesario Then
106             Call WriteConsoleMsg(Userindex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         Select Case ObjData(.Invent.BarcoObjIndex).Subtipo
        
                Case 2  'Galera

110                 If .clase <> eClass.Assasin And _
                        .clase <> eClass.Pirat And _
                        .clase <> eClass.Bandit And _
                        .clase <> eClass.Cleric And _
                        .clase <> eClass.Thief And _
                        .clase <> eClass.Paladin Then
                
112                     Call WriteConsoleMsg(Userindex, "Solo los Piratas, Asesinos, Bandidos, Clerigos, Bandidos y Paladines pueden usar Galera!!", FontTypeNames.FONTTYPE_INFO)
                    
                        Exit Sub
                
                    End If
                
114             Case 3  'Galeón
            
116                 If .clase <> eClass.Thief And .clase <> eClass.Pirat Then
118                     Call WriteConsoleMsg(Userindex, "Solo los Ladrones y Piratas pueden usar Galeon!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
            End Select

120         .Invent.BarcoObjIndex = .Invent.Object(slot).ObjIndex
122         .Invent.BarcoSlot = slot

124         If .flags.Montado > 0 Then
126             Call DoMontar(Userindex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
            End If

128         If .flags.Navegando = 0 Then

130             Call WriteNadarToggle(Userindex, IIf(Barco.Ropaje = iTraje, True, False))

132             If Barco.Ropaje <> iTraje Then
134                 .Char.Head = 0
136                 .Char.CascoAnim = NingunCasco

                End If
    
138             If .flags.Muerto = 0 Then

                    '(Nacho)
140                 If .Faccion.ArmadaReal = 1 Then
142                     If Barco.Ropaje = iTraje Then .Char.Body = iTraje
144                     If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCiuda
146                     If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCiuda
148                     If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCiuda
150                 ElseIf .Faccion.FuerzasCaos = 1 Then

152                     If Barco.Ropaje = iTraje Then .Char.Body = iTraje
154                     If Barco.Ropaje = iBarca Then .Char.Body = iBarcaPk
156                     If Barco.Ropaje = iGalera Then .Char.Body = iGaleraPk
158                     If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonPk
                    Else

160                     If Barco.Ropaje = iTraje Then .Char.Body = iTraje
162                     If Barco.Ropaje = iBarca Then .Char.Body = iBarca
164                     If Barco.Ropaje = iGalera Then .Char.Body = iGalera
166                     If Barco.Ropaje = iGaleon Then .Char.Body = iGaleon

                    End If

                Else

168                 If Barco.Ropaje = iTraje Then
170                     .Char.Body = iRopaBuceoMuerto
                    Else
172                     .Char.Body = iFragataFantasmal

                    End If

174                 .Char.Head = iCabezaMuerto

                End If
    
176             .Char.ShieldAnim = NingunEscudo
178             .Char.WeaponAnim = NingunArma
                '.Char.CascoAnim = NingunCasco
180             .flags.Navegando = 1
    
182             .Char.speeding = Barco.Velocidad
    
            Else

184             Call WriteNadarToggle(Userindex, False)

186             .Char.speeding = VelocidadNormal
    
188             .flags.Navegando = 0
    
190             If .flags.Muerto = 0 Then
192                 .Char.Head = .OrigChar.Head
        
194                 If .Invent.ArmourEqpObjIndex > 0 Then
196                     .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
            
                    Else
198                     Call DarCuerpoDesnudo(Userindex)

                    End If
        
200                 If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

202                 If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

204                 If .Invent.NudilloObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.NudilloObjIndex).WeaponAnim

206                 If .Invent.HerramientaEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.HerramientaEqpObjIndex).WeaponAnim

208                 If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                Else
210                 .Char.Body = iCuerpoMuerto
212                 .Char.Head = iCabezaMuerto
214                 .Char.ShieldAnim = NingunEscudo
216                 .Char.WeaponAnim = NingunArma
218                 .Char.CascoAnim = NingunCasco

                End If

            End If

220         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))

            'Call WriteVelocidadToggle(UserIndex)
    
222         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
224         Call WriteNavigateToggle(Userindex)
226         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, .Pos.X, .Pos.Y))
    
        End With
        
        Exit Sub

DoNavega_Err:
228     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoNavega", Erl)

230     Resume Next
        
End Sub

Public Sub DoReNavega(ByVal Userindex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoReNavega_Err
        

        Dim ModNave As Long

100     If UserList(Userindex).Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
102         Call WriteConsoleMsg(Userindex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(Userindex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

106     UserList(Userindex).Invent.BarcoObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
108     UserList(Userindex).Invent.BarcoSlot = slot

110     If UserList(Userindex).flags.Montado > 0 Then
112         Call DoMontar(Userindex, ObjData(UserList(Userindex).Invent.MonturaObjIndex), UserList(Userindex).Invent.MonturaSlot)

        End If

114     If Barco.Ropaje = iTraje Then
116         Call WriteNadarToggle(Userindex, True)
        Else
118         Call WriteNadarToggle(Userindex, False)

        End If
    
120     If Barco.Ropaje <> iTraje Then
122         UserList(Userindex).Char.Head = 0
        Else
124         UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head

        End If
    
126     If UserList(Userindex).flags.Muerto = 0 Then

            '(Nacho)
128         If UserList(Userindex).Faccion.ArmadaReal = 1 Then
130             If Barco.Ropaje = iTraje Then UserList(Userindex).Char.Body = iTraje
132             If Barco.Ropaje = iBarca Then UserList(Userindex).Char.Body = iBarcaCiuda
134             If Barco.Ropaje = iGalera Then UserList(Userindex).Char.Body = iGaleraCiuda
136             If Barco.Ropaje = iGaleon Then UserList(Userindex).Char.Body = iGaleonCiuda
138         ElseIf UserList(Userindex).Faccion.FuerzasCaos = 1 Then

140             If Barco.Ropaje = iTraje Then UserList(Userindex).Char.Body = iTraje
142             If Barco.Ropaje = iBarca Then UserList(Userindex).Char.Body = iBarcaPk
144             If Barco.Ropaje = iGalera Then UserList(Userindex).Char.Body = iGaleraPk
146             If Barco.Ropaje = iGaleon Then UserList(Userindex).Char.Body = iGaleonPk
            Else

148             If Barco.Ropaje = iTraje Then UserList(Userindex).Char.Body = iTraje
150             If Barco.Ropaje = iBarca Then UserList(Userindex).Char.Body = iBarca
152             If Barco.Ropaje = iGalera Then UserList(Userindex).Char.Body = iGalera
154             If Barco.Ropaje = iGaleon Then UserList(Userindex).Char.Body = iGaleon

            End If

        Else
156         UserList(Userindex).Char.Body = iFragataFantasmal

        End If
    
158     UserList(Userindex).Char.ShieldAnim = NingunEscudo
160     UserList(Userindex).Char.WeaponAnim = NingunArma
162     UserList(Userindex).Char.CascoAnim = NingunCasco
164     UserList(Userindex).flags.Navegando = 1
    
166     UserList(Userindex).Char.speeding = Barco.Velocidad

168     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(UserList(Userindex).Char.CharIndex, UserList(Userindex).Char.speeding))

        '
        'Call WriteVelocidadToggle(UserIndex)
    
170     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
172     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

        
        Exit Sub

DoReNavega_Err:
174     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoReNavega", Erl)
176     Resume Next
        
End Sub

Public Sub FundirMineral(ByVal Userindex As Integer)
        
        On Error GoTo FundirMineral_Err
        

100     If UserList(Userindex).flags.TargetObjInvIndex > 0 Then

            Dim SkillRequerido As Integer
102         SkillRequerido = UserList(Userindex).Stats.UserSkills(eSkill.Mineria) * ModFundirMineral(UserList(Userindex).clase)
   
104         If ObjData(UserList(Userindex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
                ObjData(UserList(Userindex).flags.TargetObjInvIndex).MinSkill <= SkillRequerido Then
            
106             Call DoLingotes(Userindex)
        
            Else
108             Call WriteConsoleMsg(Userindex, "No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas " & SkillRequerido & " puntos en minería.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

FundirMineral_Err:
110     Call RegistrarError(Err.Number, Err.description, "Trabajo.FundirMineral", Erl)
112     Resume Next
        
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal Userindex As Integer) As Boolean
        'Call LogTarea("Sub TieneObjetos")
        
        On Error GoTo TieneObjetos_Err
        

        Dim i     As Long

        Dim Total As Long

100     For i = 1 To UserList(Userindex).CurrentInventorySlots

102         If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
104             Total = Total + UserList(Userindex).Invent.Object(i).Amount

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

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal Userindex As Integer) As Boolean
        'Call LogTarea("Sub QuitarObjetos")
        
        On Error GoTo QuitarObjetos_Err
        

        Dim i As Long

100     For i = 1 To UserList(Userindex).CurrentInventorySlots
102         Debug.Print i

104         If UserList(Userindex).Invent.Object(i).ObjIndex = ItemIndex Then
106             Debug.Print UserList(Userindex).name
        
108             If UserList(Userindex).Invent.Object(i).Equipped Then
110                 Call Desequipar(Userindex, i)

                End If
        
112             UserList(Userindex).Invent.Object(i).Amount = UserList(Userindex).Invent.Object(i).Amount - cant

114             If (UserList(Userindex).Invent.Object(i).Amount <= 0) Then
116                 cant = Abs(UserList(Userindex).Invent.Object(i).Amount)
118                 UserList(Userindex).Invent.Object(i).Amount = 0
120                 UserList(Userindex).Invent.Object(i).ObjIndex = 0
                Else
122                 cant = 0

                End If
        
124             Call UpdateUserInv(False, Userindex, i)
        
126             If (cant = 0) Then
128                 QuitarObjetos = True
                    Exit Function

                End If

            End If

130     Next i

        
        Exit Function

QuitarObjetos_Err:
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.QuitarObjetos", Erl)
134     Resume Next
        
End Function

Sub HerreroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex)
102     If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex)
104     If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex)

        
        Exit Sub

HerreroQuitarMateriales_Err:
106     Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroQuitarMateriales", Erl)
108     Resume Next
        
End Sub

Sub CarpinteroQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, Userindex)

        
        Exit Sub

CarpinteroQuitarMateriales_Err:
102     Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroQuitarMateriales", Erl)
104     Resume Next
        
End Sub

Sub AlquimistaQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, Userindex)

        
        Exit Sub

AlquimistaQuitarMateriales_Err:
102     Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaQuitarMateriales", Erl)
104     Resume Next
        
End Sub

Sub SastreQuitarMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreQuitarMateriales_Err
        

100     If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, Userindex)
102     If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, Userindex)
104     If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, Userindex)

        
        Exit Sub

SastreQuitarMateriales_Err:
106     Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreQuitarMateriales", Erl)
108     Resume Next
        
End Sub

Function CarpinteroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo CarpinteroTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).Madera > 0 Then
102         If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, Userindex) Then
104             Call WriteConsoleMsg(Userindex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
106             CarpinteroTieneMateriales = False
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If
    
110     CarpinteroTieneMateriales = True

        
        Exit Function

CarpinteroTieneMateriales_Err:
112     Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroTieneMateriales", Erl)
114     Resume Next
        
End Function

Function AlquimistaTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo AlquimistaTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).Raices > 0 Then
102         If Not TieneObjetos(Raices, ObjData(ItemIndex).Raices, Userindex) Then
104             Call WriteConsoleMsg(Userindex, "No tenes suficientes raices.", FontTypeNames.FONTTYPE_INFO)
106             AlquimistaTieneMateriales = False
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If
    
110     AlquimistaTieneMateriales = True

        
        Exit Function

AlquimistaTieneMateriales_Err:
112     Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaTieneMateriales", Erl)
114     Resume Next
        
End Function

Function SastreTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo SastreTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).PielLobo > 0 Then
102         If Not TieneObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, Userindex) Then
104             Call WriteConsoleMsg(Userindex, "No tenes suficientes pieles de lobo.", FontTypeNames.FONTTYPE_INFO)
106             SastreTieneMateriales = False
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If
    
110     If ObjData(ItemIndex).PielOsoPardo > 0 Then
112         If Not TieneObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, Userindex) Then
114             Call WriteConsoleMsg(Userindex, "No tenes suficientes pieles de oso pardo.", FontTypeNames.FONTTYPE_INFO)
116             SastreTieneMateriales = False
118             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If
    
120     If ObjData(ItemIndex).PielOsoPolaR > 0 Then
122         If Not TieneObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, Userindex) Then
124             Call WriteConsoleMsg(Userindex, "No tenes suficientes pieles de oso polar.", FontTypeNames.FONTTYPE_INFO)
126             SastreTieneMateriales = False
128             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If
    
130     SastreTieneMateriales = True

        
        Exit Function

SastreTieneMateriales_Err:
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreTieneMateriales", Erl)
134     Resume Next
        
End Function

Function HerreroTieneMateriales(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo HerreroTieneMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then
102         If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, Userindex) Then
104             Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
106             HerreroTieneMateriales = False
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If

110     If ObjData(ItemIndex).LingP > 0 Then
112         If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, Userindex) Then
114             Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
116             HerreroTieneMateriales = False
118             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If

120     If ObjData(ItemIndex).LingO > 0 Then
122         If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, Userindex) Then
124             Call WriteConsoleMsg(Userindex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
126             HerreroTieneMateriales = False
128             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Function

            End If

        End If

130     HerreroTieneMateriales = True

        
        Exit Function

HerreroTieneMateriales_Err:
132     Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroTieneMateriales", Erl)
134     Resume Next
        
End Function

Public Function PuedeConstruir(ByVal Userindex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruir_Err
        
100     PuedeConstruir = HerreroTieneMateriales(Userindex, ItemIndex) And UserList(Userindex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkHerreria

        
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

Public Sub HerreroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroConstruirItem_Err
        

100     If PuedeConstruir(Userindex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
102         Call HerreroQuitarMateriales(Userindex, ItemIndex)
104         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - 2
106         Call WriteUpdateSta(Userindex)
            ' AGREGAR FX
    
108         If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el arma!", FontTypeNames.FONTTYPE_INFO)
110             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, " 1", 5))
112         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el escudo!", FontTypeNames.FONTTYPE_INFO)
114             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, " 1", 5))
116         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el casco!", FontTypeNames.FONTTYPE_INFO)
118             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, " 1", 5))
120         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
                'Call WriteConsoleMsg(UserIndex, "Has construido la armadura!", FontTypeNames.FONTTYPE_INFO)
122             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, " 1", 5))

            End If

            Dim MiObj As obj

124         MiObj.Amount = 1
126         MiObj.ObjIndex = ItemIndex

128         If Not MeterItemEnInventario(Userindex, MiObj) Then
130             Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            'End If
    
132         Call SubirSkill(Userindex, eSkill.Herreria)
134         Call UpdateUserInv(True, Userindex, 0)
136         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

138         UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

        End If

        
        Exit Sub

HerreroConstruirItem_Err:
140     Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroConstruirItem", Erl)
142     Resume Next
        
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

Public Sub CarpinteroConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroConstruirItem_Err
        

100     If CarpinteroTieneMateriales(Userindex, ItemIndex) _
                And UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria _
                And PuedeConstruirCarpintero(ItemIndex) _
                And ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).Subtipo = 5 Then
    
102         If UserList(Userindex).Stats.MinSta > 2 Then
104             Call QuitarSta(Userindex, 2)
        
            Else
106             Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para trabajar.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Sub

            End If
    
110         Call CarpinteroQuitarMateriales(Userindex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido un objeto!", FontTypeNames.FONTTYPE_INFO)
            'Call WriteOroOverHead(UserIndex, 1, UserList(UserIndex).Char.CharIndex)
112         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, " 1", 5))
    
            Dim MiObj As obj

114         MiObj.Amount = 1
116         MiObj.ObjIndex = ItemIndex

118         If Not MeterItemEnInventario(Userindex, MiObj) Then
120             Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            ' End If
    
122         Call SubirSkill(Userindex, eSkill.Carpinteria)
            'Call UpdateUserInv(True, UserIndex, 0)
124         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

126         UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

        End If

        
        Exit Sub

CarpinteroConstruirItem_Err:
128     Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroConstruirItem", Erl)
130     Resume Next
        
End Sub

Public Sub AlquimistaConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaConstruirItem_Err
        

        Rem Debug.Print UserList(UserIndex).Invent.HerramientaEqpObjIndex

100     If Not UserList(Userindex).Stats.MinSta > 0 Then
102         Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If AlquimistaTieneMateriales(Userindex, ItemIndex) _
                And UserList(Userindex).Stats.UserSkills(eSkill.Alquimia) >= ObjData(ItemIndex).SkPociones _
                And PuedeConstruirAlquimista(ItemIndex) _
                And ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).Subtipo = 4 Then
        
106         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - 25
108         Call WriteUpdateSta(Userindex)
    
110         Call AlquimistaQuitarMateriales(Userindex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
112         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(117, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
    
            Dim MiObj As obj

114         MiObj.Amount = 1
116         MiObj.ObjIndex = ItemIndex

118         If Not MeterItemEnInventario(Userindex, MiObj) Then
120             Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ''If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            'End If
    
122         Call SubirSkill(Userindex, eSkill.Alquimia)
124         Call UpdateUserInv(True, Userindex, 0)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

126         UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

        End If

        
        Exit Sub

AlquimistaConstruirItem_Err:
128     Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaConstruirItem", Erl)
130     Resume Next
        
End Sub

Public Sub SastreConstruirItem(ByVal Userindex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreConstruirItem_Err
        

100     If Not UserList(Userindex).Stats.MinSta > 0 Then
102         Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If SastreTieneMateriales(Userindex, ItemIndex) _
                And UserList(Userindex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkMAGOria _
                And PuedeConstruirSastre(ItemIndex) _
                And ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).OBJType = eOBJType.otHerramientas _
                And ObjData(UserList(Userindex).Invent.HerramientaEqpObjIndex).Subtipo = 9 Then
        
106         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - 2
        
108         Call WriteUpdateSta(Userindex)
    
110         Call SastreQuitarMateriales(Userindex, ItemIndex)
    
            ' If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
            ' Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
112         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, " 1", 5))
            ' UserList(UserIndex).flags.UltimoMensaje = 9
            ' End If
        
114         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(63, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
    
            Dim MiObj As obj

116         MiObj.Amount = 1
118         MiObj.ObjIndex = ItemIndex

120         If Not MeterItemEnInventario(Userindex, MiObj) Then
122             Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)
            End If
    
124         Call SubirSkill(Userindex, eSkill.Herreria)
126         Call UpdateUserInv(True, Userindex, 0)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

128         UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1

        End If
    
        
        Exit Sub

SastreConstruirItem_Err:
130     Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreConstruirItem", Erl)
132     Resume Next
        
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

Public Sub DoLingotes(ByVal Userindex As Integer)
        '    Call LogTarea("Sub DoLingotes")
        
        On Error GoTo DoLingotes_Err
        

100     If UserList(Userindex).Stats.MinSta > 5 Then
102         Call QuitarSta(Userindex, 5)
    
        Else
        
104         Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
106         Call WriteMacroTrabajoToggle(Userindex, False)
            Exit Sub

        End If

        Dim slot As Integer
        Dim obji As Integer

108     slot = UserList(Userindex).flags.TargetObjInvSlot
110     obji = UserList(Userindex).Invent.Object(slot).ObjIndex
    
112     Dim cant As Byte: cant = RandomNumber(1, 3)
    
        Dim necesarios As Integer

114     necesarios = MineralesParaLingote(obji, cant)
    
116     If UserList(Userindex).Invent.Object(slot).Amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
118         Call WriteConsoleMsg(Userindex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
120         Call WriteMacroTrabajoToggle(Userindex, False)
            Exit Sub

        End If
    
122     UserList(Userindex).Invent.Object(slot).Amount = UserList(Userindex).Invent.Object(slot).Amount - MineralesParaLingote(obji, cant)

124     If UserList(Userindex).Invent.Object(slot).Amount < 1 Then
126         UserList(Userindex).Invent.Object(slot).Amount = 0
128         UserList(Userindex).Invent.Object(slot).ObjIndex = 0

        End If
    
        Dim nPos  As WorldPos

        Dim MiObj As obj

130     MiObj.Amount = cant
132     MiObj.ObjIndex = ObjData(UserList(Userindex).flags.TargetObjInvIndex).LingoteIndex

134     If Not MeterItemEnInventario(Userindex, MiObj) Then
136         Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

        End If

138     Call UpdateUserInv(False, Userindex, slot)
    
140     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, cant, 5))
        'If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
        '  Call WriteConsoleMsg(UserIndex, "¡Has obtenido lingotes!", FontTypeNames.FONTTYPE_INFO)
            
        '  UserList(UserIndex).flags.UltimoMensaje = 5
        'End If
    
142     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(117, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
    
144     Call SubirSkill(Userindex, eSkill.Herreria)
  
146     UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando + 1
        
148     If UserList(Userindex).Counters.Trabajando = 1 And Not UserList(Userindex).flags.UsandoMacro Then
150         Call WriteMacroTrabajoToggle(Userindex, True)

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

Sub DoAdminInvisible(ByVal Userindex As Integer)
        
        On Error GoTo DoAdminInvisible_Err
    
100     With UserList(Userindex)
    
102         If .flags.AdminInvisible = 0 Then
                
104             .flags.AdminInvisible = 1
106             .flags.invisible = 1
108             .flags.Oculto = 1
            
                '.flags.OldBody = .Char.Body
                '.flags.OldHead = .Char.Head
            
                '.Char.Body = 0
                '.Char.Head = 0
            
110             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            
112             Call SendData(SendTarget.ToPCAreaButGMs, Userindex, PrepareMessageCharacterRemove(.Char.CharIndex, True))
            
            Else
        
114             .flags.AdminInvisible = 0
116             .flags.invisible = 0
118             .flags.Oculto = 0
120             .Counters.TiempoOculto = 0
            
                '.Char.Body = .flags.OldBody
                '.Char.Head = .flags.OldHead
            
122             Call MakeUserChar(True, 0, Userindex, .Pos.Map, .Pos.X, .Pos.Y, 1)
124             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            
            End If
        
            'Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    
        End With

        Exit Sub

DoAdminInvisible_Err:
126     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoAdminInvisible", Erl)

128     Resume Next
        
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
        
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
112         Call WriteConsoleMsg(Userindex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     If Distancia(posMadera, UserList(Userindex).Pos) > 2 Then
116         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

118     If UserList(Userindex).flags.Muerto = 1 Then
120         Call WriteConsoleMsg(Userindex, "No podés hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

122     If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
124         Call WriteConsoleMsg(Userindex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

126     If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
128         Suerte = 3
130     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
132         Suerte = 2
134     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
136         Suerte = 1

        End If

138     exito = RandomNumber(1, Suerte)

140     If exito = 1 Then
142         obj.ObjIndex = FOGATA_APAG
144         obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
146         Call WriteConsoleMsg(Userindex, "Has hecho " & obj.Amount & " ramitas.", FontTypeNames.FONTTYPE_INFO)
    
148         Call MakeObj(obj, Map, X, Y)
    
            'Seteamos la fogata como el nuevo TargetObj del user
150         UserList(Userindex).flags.TargetObj = FOGATA_APAG
        Else

            '[CDT 17-02-2004]
152         If Not UserList(Userindex).flags.UltimoMensaje = 10 Then
154             Call WriteConsoleMsg(Userindex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
156             UserList(Userindex).flags.UltimoMensaje = 10

            End If

            '[/CDT]
        End If

158     Call SubirSkill(Userindex, Supervivencia)

        
        Exit Sub

TratarDeHacerFogata_Err:
160     Call RegistrarError(Err.Number, Err.description, "Trabajo.TratarDeHacerFogata", Erl)
162     Resume Next
        
End Sub

Public Sub DoPescar(ByVal Userindex As Integer, Optional ByVal RedDePesca As Boolean = False, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte       As Integer
        Dim res          As Integer
        Dim RestaStamina As Byte

100     RestaStamina = IIf(RedDePesca, 2, 1)
    
102     With UserList(Userindex)
    
104         If .Stats.MinSta > RestaStamina Then
106             Call QuitarSta(Userindex, RestaStamina)
        
            Else
            
108             Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
            
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para pescar.", FontTypeNames.FONTTYPE_INFO)
            
110             Call WriteMacroTrabajoToggle(Userindex, False)
            
                Exit Sub

            End If

            Dim Skill As Integer

112         Skill = .Stats.UserSkills(eSkill.Pescar)
        
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
116         res = RandomNumber(1, Suerte)
    
118         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageArmaMov(.Char.CharIndex))

120         If res < 6 Then

                Dim nPos  As WorldPos
                Dim MiObj As obj
        
122             MiObj.Amount = IIf(RedDePesca, RandomNumber(2, 5), IIf(ObjetoDorado, RandomNumber(1, 3), 1)) * RecoleccionMult
124             MiObj.ObjIndex = ObtenerPezRandom(2)
        
126             If MiObj.ObjIndex = 0 Then Exit Sub
        
128             If Not MeterItemEnInventario(Userindex, MiObj) Then
130                 Call TirarItemAlPiso(.Pos, MiObj)
                End If

132             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(.Pos.X, .Pos.Y, MiObj.Amount, 5))
        
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
                
144                     If Not MeterItemEnInventario(Userindex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                    
                        ' Le mandamos un mensaje
146                     Call WriteConsoleMsg(Userindex, "¡Has conseguido " & ObjData(EspecialesPesca(i).ObjIndex).name & "!", FontTypeNames.FONTTYPE_INFO)

                        ' TODO: Sonido ?
                        'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                    End If

                Next

            End If
    
148         Call SubirSkill(Userindex, eSkill.Pescar)
    
150         .Counters.Trabajando = .Counters.Trabajando + 1
    
            'Ladder 06/07/14 Activamos el macro de trabajo
152         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
154             Call WriteMacroTrabajoToggle(Userindex, True)
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

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)
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

100     If Not MapInfo(UserList(victimaindex).Pos.Map).Seguro Then Exit Sub
    
        'If UserList(VictimaIndex).flags.EnConsulta Then
            'Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
            'Exit Sub
        'End If
    
102     With UserList(LadrOnIndex)
    
104         If .flags.Seguro Then
        
106             If Status(LadrOnIndex) = 1 Then
108                 Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            Else

110             If .Faccion.ArmadaReal = 1 Then
            
112                 If Status(victimaindex) = 1 Then
114                     Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejercito real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub

                    End If

                End If

            End If
        
            ' Caos robando a caos?
116         If UserList(victimaindex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
118             Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legion oscura.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        
120         If TriggerZonaPelea(LadrOnIndex, victimaindex) <> TRIGGER6_AUSENTE Then Exit Sub
        
            ' Tiene energia?
122         If .Stats.MinSta < 15 Then
        
124             If .genero = eGenero.Hombre Then
126                 Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
                
                Else
128                 Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
                Exit Sub

            End If
        
130         If .GuildIndex > 0 Then
        
132             If .flags.SeguroClan Then
            
134                 If .GuildIndex = UserList(victimaindex).GuildIndex Then
136                     Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

                End If

            End If

138         If .Grupo.EnGrupo > 0 Then
        
140             If .GuildIndex = UserList(victimaindex).GuildIndex Then
142                 Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If

144         If .Grupo.EnGrupo = True Then

                Dim i As Byte
146             For i = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros

148                 If UserList(.Grupo.Lider).Grupo.Miembros(i) = victimaindex Then
150                     Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

152             Next i

            End If
        
            ' Quito energia
154         Call QuitarSta(LadrOnIndex, 15)
                
156         If UserList(victimaindex).flags.Privilegios And PlayerType.user Then
            
                Dim Suerte     As Integer
                Dim res        As Integer
                Dim RobarSkill As Byte
            
158             RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
160             If RobarSkill <= 10 Then
162                 Suerte = 35
                
164             ElseIf RobarSkill <= 20 Then
166                 Suerte = 30
                
168             ElseIf RobarSkill <= 30 Then
170                 Suerte = 28
                
172             ElseIf RobarSkill <= 40 Then
174                 Suerte = 24
                
176             ElseIf RobarSkill <= 50 Then
178                 Suerte = 22
                
180             ElseIf RobarSkill <= 60 Then
182                 Suerte = 20
                
184             ElseIf RobarSkill <= 70 Then
186                 Suerte = 18
                
188             ElseIf RobarSkill <= 80 Then
190                 Suerte = 15
                
192             ElseIf RobarSkill <= 90 Then
194                 Suerte = 10
                
196             ElseIf RobarSkill < 100 Then
198                 Suerte = 7
                
                Else
200                 Suerte = 5
                
                End If
            
202             res = RandomNumber(1, Suerte)
                
204             If res < 3 Then 'Exito robo
                
206                 If UserList(victimaindex).flags.Comerciando Then
208                     OtroUserIndex = UserList(victimaindex).ComUsu.DestUsu
                        
210                     If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
212                         Call WriteConsoleMsg(victimaindex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_TALK)
214                         Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
216                         Call LimpiarComercioSeguro(victimaindex)

                        End If

                    End If
               
218                 If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    
220                     If TieneObjetosRobables(victimaindex) Then
222                         Call RobarObjeto(LadrOnIndex, victimaindex)
                        Else
224                         Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else 'Roba oro

226                     If UserList(victimaindex).Stats.GLD > 0 Then

                            Dim n As Long
                        
228                         If .clase = eClass.Thief Then n = RandomNumber(1, 100)

230                         If n > UserList(victimaindex).Stats.GLD Then n = UserList(victimaindex).Stats.GLD
                        
232                         UserList(victimaindex).Stats.GLD = UserList(victimaindex).Stats.GLD - n
                        
234                         .Stats.GLD = .Stats.GLD + n

236                         If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO
                        
238                         Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(victimaindex).name, FontTypeNames.FONTTYPE_INFO)
240                         Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
242                         Call WriteUpdateGold(victimaindex) 'Le actualizamos la billetera a la victima
                        Else
244                         Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                
246                 Call SubirSkill(LadrOnIndex, eSkill.Robar)
            
                Else
248                 Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
250                 Call WriteConsoleMsg(victimaindex, "" & .name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
252                 Call SubirSkill(LadrOnIndex, eSkill.Robar)

                End If
            
254             If Status(LadrOnIndex) = 1 Then Call VolverCriminal(LadrOnIndex)
        
256             If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)
            
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
258     Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

Public Function ObjEsRobable(ByVal victimaindex As Integer, ByVal slot As Integer) As Boolean
        ' Agregué los barcos
        ' Esta funcion determina qué objetos son robables.
        
        On Error GoTo ObjEsRobable_Err
        

        Dim OI As Integer

100     OI = UserList(victimaindex).Invent.Object(slot).ObjIndex

102     ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(victimaindex).Invent.Object(slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).donador = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otRunas And ObjData(OI).Instransferible = 0 And ObjData(OI).OBJType <> eOBJType.otMonturas

        
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

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
        '***************************************************

        Dim flag As Boolean
        Dim i    As Integer

100     flag = False

102     With UserList(victimaindex)

104         If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
106             i = 1

108             Do While Not flag And i <= .CurrentInventorySlots

                    'Hay objeto en este slot?
110                 If .Invent.Object(i).ObjIndex > 0 Then
                
112                     If ObjEsRobable(victimaindex, i) Then
                    
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
                
124                     If ObjEsRobable(victimaindex, i) Then
                    
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
144                 Call QuitarUserInvItem(victimaindex, CByte(i), 1)

                End If
                
146             Call UpdateUserInv(False, victimaindex, CByte(i))
                    
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
160         Call CancelExit(victimaindex)

        End With

End Sub

Public Sub DoApuñalar(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
        
        On Error GoTo DoApuñalar_Err
        

        '***************************************************
        'Autor: Nacho (Integer) & Unknown (orginal version)
        'Last Modification: 04/17/08 - (NicoNZ)
        'Simplifique la cuenta que hacia para sacar la suerte
        'y arregle la cuenta que hacia para sacar el daño
        '***************************************************
        Dim Suerte As Integer

        Dim Skill  As Integer
    
100     Skill = UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar)
    
102     Select Case UserList(Userindex).clase

            Case eClass.Assasin '35
104             Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 4.25)
        
106         Case eClass.Cleric, eClass.Paladin, eClass.Pirat ' 15
108             Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
        
110         Case eClass.Bard, eClass.Druid '13
112             Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
        
114         Case Else '8
116             Suerte = Int(0.0361 * Skill + 4.39)

        End Select
    
118     If RandomNumber(0, 70) < Suerte Then
120         If VictimUserIndex <> 0 Then
122             daño = daño * 1.5
            
124             UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño

126             Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageEfectOverHead("¡" & daño & "!", UserList(VictimUserIndex).Char.CharIndex, vbYellow))

128             If UserList(Userindex).ChatCombate = 1 Then
                    'Call WriteEfectOverHead(UserIndex, daño, UserList(UserIndex).Char.CharIndex) 'LADDER 21.11.08
130                 Call WriteConsoleMsg(Userindex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

                End If

132             If UserList(VictimUserIndex).ChatCombate = 1 Then
134                 Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(Userindex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
            
            Else
136             daño = daño * 2

138             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño

140             If UserList(Userindex).ChatCombate = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 1.5), FontTypeNames.FONTTYPE_FIGHT)
142                 Call WriteLocaleMsg(Userindex, "212", FontTypeNames.FONTTYPE_FIGHT, daño)

                End If
            
144             Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageEfectOverHead("¡" & daño & "!", Npclist(VictimNpcIndex).Char.CharIndex, vbYellow))

                '[Alejo]
146             Call CalcularDarExp(Userindex, VictimNpcIndex, daño)

            End If

        Else

148         If UserList(Userindex).ChatCombate = 1 Then
150             Call WriteConsoleMsg(Userindex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If
    
152     Call SubirSkill(Userindex, Apuñalar)

        
        Exit Sub

DoApuñalar_Err:
154     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoApuñalar", Erl)
156     Resume Next
        
End Sub

Public Sub DoGolpeCritico(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
        
        On Error GoTo DoGolpeCritico_Err

        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 28/01/2007
        '***************************************************
        Dim Suerte As Integer
        Dim Skill  As Integer
        
100     With UserList(Userindex)
        
102         If .clase <> eClass.Bandit Then Exit Sub
104         If .Invent.WeaponEqpSlot = 0 Then Exit Sub
106         If ObjData(.Invent.WeaponEqpObjIndex).name <> "Espada Vikinga" Then Exit Sub

108         Skill = .Stats.UserSkills(eSkill.Wrestling)
        
        End With

110     Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

112     If RandomNumber(0, 100) < Suerte Then
114         daño = Int(daño * 0.5)

116         If VictimUserIndex <> 0 Then
118             UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
120             Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
122             Call WriteConsoleMsg(VictimUserIndex, UserList(Userindex).name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
                
            Else
124             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
126             Call WriteConsoleMsg(Userindex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
128             Call CalcularDarExp(Userindex, VictimNpcIndex, daño)

            End If

        End If
        
        Exit Sub

DoGolpeCritico_Err:
130     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoGolpeCritico", Erl)

132     Resume Next
        
End Sub

Public Sub QuitarSta(ByVal Userindex As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarSta_Err
        
100     UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Cantidad

102     If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
104     If UserList(Userindex).Stats.MinSta = 0 Then Exit Sub
106     Call WriteUpdateSta(Userindex)

        
        Exit Sub

QuitarSta_Err:
108     Call RegistrarError(Err.Number, Err.description, "Trabajo.QuitarSta", Erl)
110     Resume Next
        
End Sub

Public Sub DoRaices(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
    
100     With UserList(Userindex)
    
102         If .Stats.MinSta > 2 Then
104             Call QuitarSta(Userindex, 2)
        
            Else
            
106             Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para obtener raices.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Sub
    
            End If
    
            Dim Skill As Integer
110             Skill = .Stats.UserSkills(eSkill.Alquimia)
        
112         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
114         res = RandomNumber(1, Suerte)
    
116         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageArmaMov(.Char.CharIndex))
    
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
    
136                 Call AgregarItemLimpiza(.Pos.Map, X, Y)
                
                End If
        
138             If Not MeterItemEnInventario(Userindex, MiObj) Then
            
140                 Call TirarItemAlPiso(.Pos, MiObj)
            
                End If
        
                'Call WriteConsoleMsg(UserIndex, "¡Has conseguido algunas raices!", FontTypeNames.FONTTYPE_INFO)
142             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(.Pos.X, .Pos.Y, MiObj.Amount, 5))
144             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(60, .Pos.X, .Pos.Y))
            Else
146             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(61, .Pos.X, .Pos.Y))
    
                '[CDT 17-02-2004]
148             If Not .flags.UltimoMensaje = 8 Then
150                 Call WriteConsoleMsg(Userindex, "¡No has obtenido raices!", FontTypeNames.FONTTYPE_INFO)
152                 .flags.UltimoMensaje = 8
    
                End If
        
                '[/CDT]
            End If
    
154         Call SubirSkill(Userindex, eSkill.Alquimia)
    
156         .Counters.Trabajando = .Counters.Trabajando + 1
    
158         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
160             Call WriteMacroTrabajoToggle(Userindex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
162     Call LogError("Error en DoRaices")

End Sub

Public Sub DoTalar(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
    
100     With UserList(Userindex)
    
                'EsfuerzoTalarLeñador = 1
102         If .Stats.MinSta > 2 Then
104             Call QuitarSta(Userindex, 2)
        
            Else
106             Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para talar.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Sub
    
            End If
    
            Dim Skill As Integer
    
110         Skill = .Stats.UserSkills(eSkill.Talar)
112         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
114         res = RandomNumber(1, Suerte)
116         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageArmaMov(.Char.CharIndex))
        
118         If res < 6 Then
    
                Dim nPos  As WorldPos
    
                Dim MiObj As obj
            
120             If .flags.TargetObj = 0 Then Exit Sub
            
122             Call ActualizarRecurso(.Pos.Map, X, Y)
124             MapData(.Pos.Map, X, Y).ObjInfo.data = GetTickCount() ' Ultimo uso
    
126             MiObj.Amount = IIf(ObjetoDorado, RandomNumber(1, 5), 1) * RecoleccionMult
128             MiObj.ObjIndex = Leña
            
130             If MiObj.Amount > MapData(.Pos.Map, X, Y).ObjInfo.Amount Then
132                 MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                End If
            
134             MapData(.Pos.Map, X, Y).ObjInfo.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount - MiObj.Amount
            
136             If Not MeterItemEnInventario(Userindex, MiObj) Then
138                 Call TirarItemAlPiso(.Pos, MiObj)
                End If
    
                'If Not .flags.UltimoMensaje = 5 Then
                ' Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
                '        .flags.UltimoMensaje = 5
                ' End If
140             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateRenderValue(.Pos.X, .Pos.Y, MiObj.Amount, 5))
142             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
            
                ' Al talar también podés dropear cosas raras (se setean desde RecursosEspeciales.dat)
                Dim i As Integer
    
                ' Por cada drop posible
144             For i = 1 To UBound(EspecialesTala)
                    ' Tiramos al azar entre 1 y la probabilidad
146                 res = RandomNumber(1, EspecialesTala(i).data)
                
                    ' Si tiene suerte y le pega
148                 If res = 1 Then
150                     MiObj.ObjIndex = EspecialesTala(i).ObjIndex
152                     MiObj.Amount = 1 ' Solo un item por vez
                    
                        'If Not MeterItemEnInventario(Userindex, MiObj) Then _
                        'Call TirarItemAlPiso(.Pos, MiObj)
    
                        ' Tiro siempre el item al piso, me parece más rolero, como que cae del árbol :P
154                     Call TirarItemAlPiso(.Pos, MiObj)
    
                        ' Oculto el mensaje porque el item cae al piso
                        'Call WriteConsoleMsg(Userindex, "¡Has conseguido " & ObjData(EspecialesTala(i).ObjIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
                        ' TODO: Sonido ?
                        'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                    End If
    
                Next
        
            Else
                '[CDT 17-02-2004]
156             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(64, .Pos.X, .Pos.Y))
    
158             If Not .flags.UltimoMensaje = 8 Then
160                 Call WriteConsoleMsg(Userindex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
162                 .flags.UltimoMensaje = 8
    
                End If
    
                '[/CDT]
            End If
        
164         Call SubirSkill(Userindex, eSkill.Talar)
        
166         .Counters.Trabajando = .Counters.Trabajando + 1
    
168         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
170             Call WriteMacroTrabajoToggle(Userindex, True)
            End If
    
        End With

        Exit Sub

ErrHandler:
172     Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
        Dim metal  As Integer

100     With UserList(Userindex)
    
            'Por Ladder 06/07/2014 Cuando la estamina llega a 0 , el macro se desactiva
102         If .Stats.MinSta > 2 Then
104             Call QuitarSta(Userindex, 2)
            Else
106             Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(Userindex, False)
                Exit Sub
    
            End If
    
            'Por Ladder 06/07/2014
    
            Dim Skill As Integer
    
110         Skill = .Stats.UserSkills(eSkill.Mineria)
112         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
114         res = RandomNumber(1, Suerte)
        
116         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageArmaMov(.Char.CharIndex))
        
118         If res <= 5 Then
    
                Dim MiObj As obj
                Dim nPos  As WorldPos
            
120             If .flags.TargetObj = 0 Then Exit Sub
            
122             Call ActualizarRecurso(.Pos.Map, X, Y)
124             MapData(.Pos.Map, X, Y).ObjInfo.data = GetTickCount() ' Ultimo uso
            
126             MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
128             MiObj.Amount = IIf(ObjetoDorado, RandomNumber(1, 6), 1) * RecoleccionMult
            
130             If MiObj.Amount > MapData(.Pos.Map, X, Y).ObjInfo.Amount Then
132                 MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                End If
            
134             MapData(.Pos.Map, X, Y).ObjInfo.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount - MiObj.Amount
        
136             If Not MeterItemEnInventario(Userindex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
            
138             Call WriteConsoleMsg(Userindex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
140             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.X, .Pos.Y))
            
                ' Al minar también puede dropear una gema
                Dim i As Integer
    
                ' Por cada drop posible
142             For i = 1 To ObjData(.flags.TargetObj).CantItem
                    ' Tiramos al azar entre 1 y la probabilidad
144                 res = RandomNumber(1, ObjData(.flags.TargetObj).Item(i).Amount)
                
                    ' Si tiene suerte y le pega
146                 If res = 1 Then
                        ' Se lo metemos al inventario (o lo tiramos al piso)
148                     MiObj.ObjIndex = ObjData(.flags.TargetObj).Item(i).ObjIndex
150                     MiObj.Amount = 1 ' Solo una gema por vez
                    
152                     If Not MeterItemEnInventario(Userindex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                        
                        ' Le mandamos un mensaje
154                     Call WriteConsoleMsg(Userindex, "¡Has conseguido " & ObjData(ObjData(.flags.TargetObj).Item(i).ObjIndex).name & "!", FontTypeNames.FONTTYPE_INFO)
                        ' TODO: Sonido de drop de gema :P
                        'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                        
                        ' Como máximo dropea una gema
                        'Exit For ' Lo saco a pedido de Haracin
                    End If
    
                Next
            
            Else
156             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(62, .Pos.X, .Pos.Y))
    
                '[CDT 17-02-2004]
158             If Not .flags.UltimoMensaje = 9 Then
                
160                 Call WriteConsoleMsg(Userindex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                
162                 .flags.UltimoMensaje = 9
    
                End If
    
                '[/CDT]
            End If
        
164         Call SubirSkill(Userindex, eSkill.Mineria)
        
166         .Counters.Trabajando = .Counters.Trabajando + 1
        
168         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
170             Call WriteMacroTrabajoToggle(Userindex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
172     Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal Userindex As Integer)
        
        On Error GoTo DoMeditar_Err
        

        Dim Suerte       As Integer
        Dim res          As Integer
        Dim cant         As Integer
        Dim MeditarSkill As Byte

100     With UserList(Userindex)

102         If .Stats.MinMAN >= .Stats.MaxMAN Then
104             .flags.Meditando = False
106             .Char.FX = 0
108             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
                Exit Sub
            End If
            
110         If (GetTickCount()) - .Counters.InicioMeditar < TIEMPO_INICIOMEDITAR Then Exit Sub
    
112         MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
            
114         If MeditarSkill <= 10 Then
116             Suerte = 35
118         ElseIf MeditarSkill <= 20 Then
120             Suerte = 30
122         ElseIf MeditarSkill <= 30 Then
124             Suerte = 28
126         ElseIf MeditarSkill <= 40 Then
128             Suerte = 24
130         ElseIf MeditarSkill <= 50 Then
132             Suerte = 22
134         ElseIf MeditarSkill <= 60 Then
136             Suerte = 20
138         ElseIf MeditarSkill <= 70 Then
140             Suerte = 18
142         ElseIf MeditarSkill <= 80 Then
144             Suerte = 15
146         ElseIf MeditarSkill <= 90 Then
148             Suerte = 10
150         ElseIf MeditarSkill < 100 Then
152             Suerte = 7
            Else
154             Suerte = 5
            End If
    
156         If .flags.RegeneracionMana = 1 Then
158             Suerte = 10
            End If
        
160         res = RandomNumber(1, Suerte)
    
162         If res = 1 Then

164             cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)

166             If cant <= 0 Then cant = 1

168             .Stats.MinMAN = .Stats.MinMAN + cant

170             If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
172             Call WriteUpdateMana(Userindex)
174             Call SubirSkill(Userindex, Meditar)

            End If

        End With

        
        Exit Sub

DoMeditar_Err:
176     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoMeditar", Erl)
178     Resume Next
        
End Sub

Public Sub DoDesequipar(ByVal Userindex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modif: 15/04/2010
        'Unequips either shield, weapon or helmet from target user.
        '***************************************************

        Dim Probabilidad   As Integer
        Dim Resultado      As Integer
        Dim WrestlingSkill As Byte
        Dim AlgoEquipado   As Boolean
    
100     With UserList(Userindex)

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
                
118                 Call WriteConsoleMsg(Userindex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
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
                
134                 Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
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
                
150                 Call WriteConsoleMsg(Userindex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
152                 If .Stats.ELV < 20 Then
154                     Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)

                    End If
                
                    Exit Sub

                End If
            
156             AlgoEquipado = True

            End If
    
158         If AlgoEquipado Then
160             Call WriteConsoleMsg(Userindex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
            Else
162             Call WriteConsoleMsg(Userindex, "No has logrado desequipar ningun item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
        End With

End Sub

Public Sub DoHurtar(ByVal Userindex As Integer, ByVal victimaindex As Integer)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modif: 03/03/2010
        'Implements the pick pocket skill of the Bandit :)
        '03/03/2010 - Pato: Solo se puede hurtar si no esta en trigger 6 :)
        '***************************************************
        Dim OtroUserIndex As Integer

100     If TriggerZonaPelea(Userindex, victimaindex) <> TRIGGER6_AUSENTE Then Exit Sub

102     If UserList(Userindex).clase <> eClass.Bandit Then Exit Sub

        Dim res As Integer

104     res = RandomNumber(1, 100)

106     If (res < 20) Then
108         If TieneObjetosRobables(victimaindex) Then
    
110             If UserList(victimaindex).flags.Comerciando Then
112                 OtroUserIndex = UserList(victimaindex).ComUsu.DestUsu
                
114                 If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
116                     Call WriteConsoleMsg(victimaindex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_WARNING)
118                     Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_WARNING)
                
120                     Call LimpiarComercioSeguro(victimaindex)

                    End If

                End If
                
122             Call RobarObjeto(Userindex, victimaindex)
124             Call WriteConsoleMsg(victimaindex, "" & UserList(Userindex).name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
            Else
126             Call WriteConsoleMsg(Userindex, UserList(victimaindex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

End Sub

Public Sub DoHandInmo(ByVal Userindex As Integer, ByVal victimaindex As Integer)

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modif: 17/02/2007
        'Implements the special Skill of the Thief
        '***************************************************
100     If UserList(victimaindex).flags.Paralizado = 1 Then Exit Sub
        
        Dim res As Integer

102     res = RandomNumber(0, 100)

104     If res < (UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
106         UserList(victimaindex).flags.Paralizado = 1
108         UserList(victimaindex).Counters.Paralisis = IntervaloParalizado / 2
        
            'UserList(VictimaIndex).flags.ParalizedByIndex = Userindex
            'UserList(VictimaIndex).flags.ParalizedBy = UserList(Userindex).name
        
110         Call WriteParalizeOK(victimaindex)
112         Call WriteConsoleMsg(Userindex, "Tu golpe ha dejado inmovil a tu oponente", FontTypeNames.FONTTYPE_FIGHT)
114         Call WriteConsoleMsg(victimaindex, "El golpe te ha dejado inmovil!", FontTypeNames.FONTTYPE_FIGHT)

        End If

End Sub

Public Sub Desarmar(ByVal Userindex As Integer, ByVal VictimIndex As Integer)
        
        On Error GoTo Desarmar_Err
        

        Dim Suerte As Integer

        Dim res    As Integer

100     If UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 10 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
102         Suerte = 35
104     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 20 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
106         Suerte = 30
108     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 30 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
110         Suerte = 28
112     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 40 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
114         Suerte = 24
116     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 50 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
118         Suerte = 22
120     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 60 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
122         Suerte = 20
124     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 70 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
126         Suerte = 18
128     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 80 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
130         Suerte = 15
132     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) <= 90 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
134         Suerte = 10
136     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) < 100 And UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
138         Suerte = 7
140     ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
142         Suerte = 5

        End If

144     res = RandomNumber(1, Suerte)

146     If res <= 2 Then
148         Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
150         Call WriteConsoleMsg(Userindex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

152         If UserList(VictimIndex).Stats.ELV < 20 Then
154             Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        

        End If

        
        Exit Sub

Desarmar_Err:
156     Call RegistrarError(Err.Number, Err.description, "Trabajo.Desarmar", Erl)
158     Resume Next
        
End Sub

Public Sub DoMontar(ByVal Userindex As Integer, ByRef Montura As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoMontar_Err
        

100     If Not CheckRazaTipo(Userindex, UserList(Userindex).Invent.Object(slot).ObjIndex) Then
102         Call WriteConsoleMsg(Userindex, "Tu raza no te permite usar esta montura.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If Not CheckClaseTipo(Userindex, UserList(Userindex).Invent.Object(slot).ObjIndex) Then
106         Call WriteConsoleMsg(Userindex, "Tu clase no te permite usar esta montura.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

108     If UserList(Userindex).Stats.UserSkills(eSkill.equitacion) < Montura.MinSkill Then
            'Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar esta montura.", FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(Userindex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        'Ladder 21/11/08
112     If UserList(Userindex).flags.Montado = 0 Then
114         If (MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger > 10) Then
116             Call WriteConsoleMsg(Userindex, "No podés montar aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

118     If UserList(Userindex).flags.Meditando Then
120         UserList(Userindex).flags.Meditando = False
122         UserList(Userindex).Char.FX = 0
124         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(UserList(Userindex).Char.CharIndex, 0))
        End If

126     If UserList(Userindex).flags.Montado = 1 Then
128         If UserList(Userindex).Invent.MonturaObjIndex > 0 Then
130             If ObjData(UserList(Userindex).Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
132                 Call UpdateUserInv(False, Userindex, UserList(Userindex).Invent.MonturaSlot)
                End If

            End If

        End If

134     UserList(Userindex).Invent.MonturaObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
136     UserList(Userindex).Invent.MonturaSlot = slot

138     If UserList(Userindex).flags.Montado = 0 Then
            
140         UserList(Userindex).Char.Body = Montura.Ropaje

            'UserList(UserIndex).Char.body = Montura.Ropaje
142         UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
144         UserList(Userindex).Char.ShieldAnim = NingunEscudo
146         UserList(Userindex).Char.WeaponAnim = NingunArma
148         UserList(Userindex).Char.CascoAnim = UserList(Userindex).Char.CascoAnim
150         UserList(Userindex).flags.Montado = 1
152         UserList(Userindex).Char.speeding = VelocidadMontura
        Else
154         UserList(Userindex).flags.Montado = 0
156         UserList(Userindex).Char.Head = UserList(Userindex).OrigChar.Head
158         UserList(Userindex).Char.speeding = VelocidadNormal

160         If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
162             UserList(Userindex).Char.Body = ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Ropaje

            Else
164             Call DarCuerpoDesnudo(Userindex)

            End If
            
166         If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then UserList(Userindex).Char.ShieldAnim = ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).ShieldAnim

168         If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then UserList(Userindex).Char.WeaponAnim = ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).WeaponAnim

170         If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then UserList(Userindex).Char.CascoAnim = ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).CascoAnim

        End If

172     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)

174     Call UpdateUserInv(False, Userindex, slot)
176     Call WriteEquiteToggle(Userindex)
178     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(UserList(Userindex).Char.CharIndex, UserList(Userindex).Char.speeding))

        
        Exit Sub

DoMontar_Err:
180     Call RegistrarError(Err.Number, Err.description, "Trabajo.DoMontar", Erl)
182     Resume Next
        
End Sub

Public Function ApuñalarFunction(ByVal Userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer) As Integer
        
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

100     Skill = UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar)

102     Select Case UserList(Userindex).clase

            Case eClass.Assasin '35
104             Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 5)
        
106             If VictimNpcIndex = 0 Then
108                 If UserList(VictimUserIndex).Char.Heading = UserList(Userindex).Char.Heading Then
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
132                 ApuñalarFunction = daño * ModicadorApuñalarClase(UserList(Userindex).clase)
            Else ' A los npcs el apuñalar es un 20% mas
134                ApuñalarFunction = (daño * ModicadorApuñalarClase(UserList(Userindex).clase)) * 1.2
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

End Function

Function FreeMascotaIndex(ByVal Userindex As Integer) As Integer

        '***************************************************
        'Author: Unknown
        'Last Modification: 02/03/09
        '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
        '***************************************************
        Dim j As Integer

100     For j = 1 To MAXMASCOTAS

102         If UserList(Userindex).MascotasType(j) = 0 Then
104             FreeMascotaIndex = j
                Exit Function

            End If

106     Next j

End Function

Sub DoDomar(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Nacho (Integer)
        'Last Modification: 01/05/2010
        '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
        '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
        '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
        '***************************************************

        On Error GoTo ErrHandler

        Dim puntosDomar      As Integer

        Dim puntosRequeridos As Integer

        Dim CanStay          As Boolean

        Dim petType          As Integer

        Dim NroPets          As Integer
    
100     If Npclist(NpcIndex).MaestroUser = Userindex Then
102         Call WriteConsoleMsg(Userindex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     With UserList(Userindex)

106         If .NroMascotas < MAXMASCOTAS Then

108             If Npclist(NpcIndex).MaestroNPC > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
110                 Call WriteConsoleMsg(Userindex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

112             If Not PuedeDomarMascota(Userindex, NpcIndex) Then
114                 Call WriteConsoleMsg(Userindex, "No puedes domar más de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

116             puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))

                ' 20% de bonificacion
118             If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
120                 puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8

                ' 11% de bonificacion
122             ElseIf .Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
124                 puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.89

                Else
126                 puntosRequeridos = Npclist(NpcIndex).flags.Domable
                End If

128             If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then

                    Dim Index As Integer

130                 .NroMascotas = .NroMascotas + 1
132                 Index = FreeMascotaIndex(Userindex)
134                 .MascotasIndex(Index) = NpcIndex
136                 .MascotasType(Index) = Npclist(NpcIndex).Numero

138                 Npclist(NpcIndex).MaestroUser = Userindex

140                 Call FollowAmo(NpcIndex)
142                 Call ReSpawnNpc(Npclist(NpcIndex))

144                 Call WriteConsoleMsg(Userindex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)

                    ' Es zona segura?
146                 If MapInfo(.Pos.Map).Seguro Then
148                     petType = Npclist(NpcIndex).Numero
150                     NroPets = .NroMascotas

152                     Call QuitarNPC(NpcIndex)

154                     .MascotasType(Index) = petType
156                     .NroMascotas = NroPets

158                     Call WriteConsoleMsg(Userindex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)
                    End If

                Else

160                 If Not .flags.UltimoMensaje = 5 Then
162                     Call WriteConsoleMsg(Userindex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
164                     .flags.UltimoMensaje = 5
                    End If

                End If

166             Call SubirSkill(Userindex, eSkill.Domar)

            Else
168             Call WriteConsoleMsg(Userindex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

ErrHandler:
170     Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal Userindex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean

        '***************************************************
        'Author: ZaMa
        'This function checks how many NPCs of the same type have
        'been tamed by the user.
        'Returns True if that amount is less than two.
        '***************************************************
        Dim i           As Long

        Dim numMascotas As Long
    
100     For i = 1 To MAXMASCOTAS

102         If UserList(Userindex).MascotasType(i) = Npclist(NpcIndex).Numero Then
104             numMascotas = numMascotas + 1

            End If

106     Next i
    
108     If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Private Function ModFundirMineral(ByVal clase As eClass) As Integer
    
100     If clase = eClass.Trabajador Then
102         ModFundirMineral = 1
        Else
104         ModFundirMineral = 3
        End If
    
End Function
