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
        

100     UserList(UserIndex).Counters.TiempoOculto = UserList(UserIndex).Counters.TiempoOculto - 1

102     If UserList(UserIndex).Counters.TiempoOculto <= 0 Then
    
            ' UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
    
104         UserList(UserIndex).Counters.TiempoOculto = 0
106         UserList(UserIndex).flags.Oculto = 0
108         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
110         Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

        End If

        Exit Sub

Errhandler:
112     Call LogError("Error en Sub DoPermanecerOculto")

        
        Exit Sub

DoPermanecerOculto_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoPermanecerOculto", Erl)
        Resume Next
        
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la fórmula y ahora anda bien.
    On Error GoTo Errhandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer

    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Ocultarse)

    Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

    res = RandomNumber(1, 100)

    If res <= Suerte Then

        UserList(UserIndex).flags.Oculto = 1
        Suerte = (-0.000001 * (100 - Skill) ^ 3)
        Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
        Suerte = Suerte + (-0.0088 * (100 - Skill))
        Suerte = Suerte + (0.9571)
        Suerte = Suerte * IntervaloOculto
    
        If UserList(UserIndex).flags.AnilloOcultismo = 1 Then
            UserList(UserIndex).Counters.TiempoOculto = Suerte * 3
        Else
            UserList(UserIndex).Counters.TiempoOculto = Suerte

        End If
  
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

        'Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
        Call WriteLocaleMsg(UserIndex, "55", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(UserIndex, Ocultarse)
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 4 Then
            'Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "57", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 4

        End If

        '[/CDT]
    End If

    UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando + 1

    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNadar(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoNadar_Err
        

        Dim ModNave As Long

100     If UserList(UserIndex).flags.Nadando = 0 Then
    
102         If UserList(UserIndex).flags.Muerto = 0 Then
                '(Nacho)
    
104             UserList(UserIndex).Char.Body = 694
                'If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGalera
                'If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleon
            Else
106             UserList(UserIndex).Char.Body = iFragataFantasmal

            End If
    
108         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
110         UserList(UserIndex).Char.WeaponAnim = NingunArma
112         UserList(UserIndex).Char.CascoAnim = NingunCasco
114         UserList(UserIndex).flags.Nadando = 1
    
        Else
    
116         UserList(UserIndex).flags.Nadando = 0
    
118         If UserList(UserIndex).flags.Muerto = 0 Then
120             UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
122             If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
124                 UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            
                Else
132                 Call DarCuerpoDesnudo(UserIndex)

                End If
        
134             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

136             If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

138             If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
            Else
140             UserList(UserIndex).Char.Body = iCuerpoMuerto
142             UserList(UserIndex).Char.Head = iCabezaMuerto
144             UserList(UserIndex).Char.ShieldAnim = NingunEscudo
146             UserList(UserIndex).Char.WeaponAnim = NingunArma
148             UserList(UserIndex).Char.CascoAnim = NingunCasco

            End If

        End If

150     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        'Call WriteNadarToggle(UserIndex)
152     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

        
        Exit Sub

DoNadar_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoNadar", Erl)
        Resume Next
        
End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoNavega_Err
        

        Dim ModNave As Long

100     If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
102         Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

106     UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
108     UserList(UserIndex).Invent.BarcoSlot = slot

110     If UserList(UserIndex).flags.Montado > 0 Then
112         Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

        End If

114     If UserList(UserIndex).flags.Navegando = 0 Then

116         If Barco.Ropaje = iTraje Then
118             Call WriteNadarToggle(UserIndex, True)
        
            Else
120             Call WriteNadarToggle(UserIndex, False)
        
            End If
    
122         If Barco.Ropaje <> iTraje Then
124             UserList(UserIndex).Char.Head = 0
126             UserList(UserIndex).Char.CascoAnim = NingunCasco

            End If
    
128         If UserList(UserIndex).flags.Muerto = 0 Then

                '(Nacho)
130             If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
132                 If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
134                 If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaCiuda
136                 If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraCiuda
138                 If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonCiuda
140             ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then

142                 If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
144                 If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaPk
146                 If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraPk
148                 If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonPk
                Else

150                 If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
152                 If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarca
154                 If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGalera
156                 If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleon

                End If

            Else

158             If Barco.Ropaje = iTraje Then
160                 UserList(UserIndex).Char.Body = iRopaBuceoMuerto
                Else
162                 UserList(UserIndex).Char.Body = iFragataFantasmal

                End If

164             UserList(UserIndex).Char.Head = iCabezaMuerto

            End If
    
166         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
168         UserList(UserIndex).Char.WeaponAnim = NingunArma
            'UserList(UserIndex).Char.CascoAnim = NingunCasco
170         UserList(UserIndex).flags.Navegando = 1
    
172         UserList(UserIndex).Char.speeding = Barco.Velocidad
    
        Else

174         Call WriteNadarToggle(UserIndex, False)

180         UserList(UserIndex).Char.speeding = VelocidadNormal
    
182         UserList(UserIndex).flags.Navegando = 0
    
184         If UserList(UserIndex).flags.Muerto = 0 Then
186             UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
188             If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
190                 UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            
                Else
198                 Call DarCuerpoDesnudo(UserIndex)

                End If
        
200             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

202             If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

204             If UserList(UserIndex).Invent.NudilloObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.NudilloObjIndex).WeaponAnim

206             If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).WeaponAnim

208             If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
            Else
210             UserList(UserIndex).Char.Body = iCuerpoMuerto
212             UserList(UserIndex).Char.Head = iCabezaMuerto
214             UserList(UserIndex).Char.ShieldAnim = NingunEscudo
216             UserList(UserIndex).Char.WeaponAnim = NingunArma
218             UserList(UserIndex).Char.CascoAnim = NingunCasco

            End If

        End If

220     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

        'Call WriteVelocidadToggle(UserIndex)
    
222     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
224     Call WriteNavigateToggle(UserIndex)
226     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

        
        Exit Sub

DoNavega_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoNavega", Erl)
        Resume Next
        
End Sub

Public Sub DoReNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)
        
        On Error GoTo DoReNavega_Err
        

        Dim ModNave As Long

100     If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
102         Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

106     UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
108     UserList(UserIndex).Invent.BarcoSlot = slot

110     If UserList(UserIndex).flags.Montado > 0 Then
112         Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

        End If

114     If Barco.Ropaje = iTraje Then
116         Call WriteNadarToggle(UserIndex, True)
        Else
118         Call WriteNadarToggle(UserIndex, False)

        End If
    
120     If Barco.Ropaje <> iTraje Then
122         UserList(UserIndex).Char.Head = 0
        Else
124         UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head

        End If
    
126     If UserList(UserIndex).flags.Muerto = 0 Then

            '(Nacho)
128         If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
130             If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
132             If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaCiuda
134             If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraCiuda
136             If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonCiuda
138         ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then

140             If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
142             If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaPk
144             If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraPk
146             If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonPk
            Else

148             If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
150             If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarca
152             If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGalera
154             If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleon

            End If

        Else
156         UserList(UserIndex).Char.Body = iFragataFantasmal

        End If
    
158     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
160     UserList(UserIndex).Char.WeaponAnim = NingunArma
162     UserList(UserIndex).Char.CascoAnim = NingunCasco
164     UserList(UserIndex).flags.Navegando = 1
    
166     UserList(UserIndex).Char.speeding = Barco.Velocidad

168     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

        '
        'Call WriteVelocidadToggle(UserIndex)
    
170     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
172     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

        
        Exit Sub

DoReNavega_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoReNavega", Erl)
        Resume Next
        
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
        
        On Error GoTo FundirMineral_Err
        

100     If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then

            Dim SkillRequerido As Integer
            SkillRequerido = UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) * ModFundirMineral(UserList(UserIndex).clase)
   
102         If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
                ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= SkillRequerido Then
            
104             Call DoLingotes(UserIndex)
        
            Else
106             Call WriteConsoleMsg(UserIndex, "No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas " & SkillRequerido & " puntos en minería.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

FundirMineral_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.FundirMineral", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.TieneObjetos", Erl)
        Resume Next
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        'Call LogTarea("Sub QuitarObjetos")
        
        On Error GoTo QuitarObjetos_Err
        

        Dim i As Long

100     For i = 1 To UserList(UserIndex).CurrentInventorySlots
102         Debug.Print i

104         If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
106             Debug.Print UserList(UserIndex).name
        
108             If UserList(UserIndex).Invent.Object(i).Equipped Then
110                 Call Desequipar(UserIndex, i)

                End If
        
112             UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - cant

114             If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
116                 cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
118                 UserList(UserIndex).Invent.Object(i).Amount = 0
120                 UserList(UserIndex).Invent.Object(i).ObjIndex = 0
                Else
122                 cant = 0

                End If
        
124             Call UpdateUserInv(False, UserIndex, i)
        
126             If (cant = 0) Then
128                 QuitarObjetos = True
                    Exit Function

                End If

            End If

130     Next i

        
        Exit Function

QuitarObjetos_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.QuitarObjetos", Erl)
        Resume Next
        
End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
102     If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
104     If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)

        
        Exit Sub

HerreroQuitarMateriales_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroQuitarMateriales", Erl)
        Resume Next
        
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)

        
        Exit Sub

CarpinteroQuitarMateriales_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroQuitarMateriales", Erl)
        Resume Next
        
End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)

        
        Exit Sub

AlquimistaQuitarMateriales_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaQuitarMateriales", Erl)
        Resume Next
        
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreQuitarMateriales_Err
        

100     If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex)
102     If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
104     If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex)

        
        Exit Sub

SastreQuitarMateriales_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreQuitarMateriales", Erl)
        Resume Next
        
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
    
110     CarpinteroTieneMateriales = True

        
        Exit Function

CarpinteroTieneMateriales_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroTieneMateriales", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaTieneMateriales", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreTieneMateriales", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroTieneMateriales", Erl)
        Resume Next
        
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruir_Err
        
100     PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkHerreria

        
        Exit Function

PuedeConstruir_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruir", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirHerreria", Erl)
        Resume Next
        
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroConstruirItem_Err
        

100     If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
102         Call HerreroQuitarMateriales(UserIndex, ItemIndex)
104         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
106         Call WriteUpdateSta(UserIndex)
            ' AGREGAR FX
    
108         If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el arma!", FontTypeNames.FONTTYPE_INFO)
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
112         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el escudo!", FontTypeNames.FONTTYPE_INFO)
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
116         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el casco!", FontTypeNames.FONTTYPE_INFO)
118             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
120         ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
                'Call WriteConsoleMsg(UserIndex, "Has construido la armadura!", FontTypeNames.FONTTYPE_INFO)
122             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))

            End If

            Dim MiObj As obj

124         MiObj.Amount = 1
126         MiObj.ObjIndex = ItemIndex

128         If Not MeterItemEnInventario(UserIndex, MiObj) Then
130             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            'End If
    
132         Call SubirSkill(UserIndex, eSkill.Herreria)
134         Call UpdateUserInv(True, UserIndex, 0)
136         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

138         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

HerreroConstruirItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.HerreroConstruirItem", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirCarpintero", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirAlquimista", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.PuedeConstruirSastre", Erl)
        Resume Next
        
End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo CarpinteroConstruirItem_Err
        

100     If CarpinteroTieneMateriales(UserIndex, ItemIndex) And _
            UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria And _
            PuedeConstruirCarpintero(ItemIndex) And _
            UserList(UserIndex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then
    
102         If UserList(UserIndex).Stats.MinSta > 2 Then
104             Call QuitarSta(UserIndex, 2)
        
            Else
106             Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para trabajar.", FontTypeNames.FONTTYPE_INFO)
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If
    
110         Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido un objeto!", FontTypeNames.FONTTYPE_INFO)
            'Call WriteOroOverHead(UserIndex, 1, UserList(UserIndex).Char.CharIndex)
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
    
            Dim MiObj As obj

114         MiObj.Amount = 1
116         MiObj.ObjIndex = ItemIndex

118         If Not MeterItemEnInventario(UserIndex, MiObj) Then
120             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            ' If ObjData(MiObj.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
            ' End If
    
122         Call SubirSkill(UserIndex, eSkill.Carpinteria)
124         Call UpdateUserInv(True, UserIndex, 0)
126         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

128         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

CarpinteroConstruirItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.CarpinteroConstruirItem", Erl)
        Resume Next
        
End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaConstruirItem_Err
        

        Rem Debug.Print UserList(UserIndex).Invent.HerramientaEqpObjIndex

100     If Not UserList(UserIndex).Stats.MinSta > 0 Then
102         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If AlquimistaTieneMateriales(UserIndex, ItemIndex) And _
            UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) >= ObjData(ItemIndex).SkPociones And _
            PuedeConstruirAlquimista(ItemIndex) And UserList(UserIndex).Invent.HerramientaEqpObjIndex = OLLA_ALQUIMIA Then
        
106         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 25
108         Call WriteUpdateSta(UserIndex)
    
110         Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
            'Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.AlquimistaConstruirItem", Erl)
        Resume Next
        
End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreConstruirItem_Err
        

100     If Not UserList(UserIndex).Stats.MinSta > 0 Then
102         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If SastreTieneMateriales(UserIndex, ItemIndex) And _
            UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkMAGOria And _
            PuedeConstruirSastre(ItemIndex) And UserList(UserIndex).Invent.HerramientaEqpObjIndex = COSTURERO Then
        
106         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        
108         Call WriteUpdateSta(UserIndex)
    
110         Call SastreQuitarMateriales(UserIndex, ItemIndex)
    
            ' If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
            ' Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
            ' UserList(UserIndex).flags.UltimoMensaje = 9
            ' End If
        
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(63, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
            Dim MiObj As obj

116         MiObj.Amount = 1
118         MiObj.ObjIndex = ItemIndex

120         If Not MeterItemEnInventario(UserIndex, MiObj) Then
122             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
    
124         Call SubirSkill(UserIndex, eSkill.Herreria)
126         Call UpdateUserInv(True, UserIndex, 0)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

128         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If
    
        
        Exit Sub

SastreConstruirItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.SastreConstruirItem", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.MineralesParaLingote", Erl)
        Resume Next
        
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
    
140     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, cant, 5))
        'If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
        '  Call WriteConsoleMsg(UserIndex, "¡Has obtenido lingotes!", FontTypeNames.FONTTYPE_INFO)
            
        '  UserList(UserIndex).flags.UltimoMensaje = 5
        'End If
    
142     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
144     Call SubirSkill(UserIndex, eSkill.Herreria)
  
146     UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
        
148     If UserList(UserIndex).Counters.Trabajando = 1 And Not UserList(UserIndex).flags.UsandoMacro Then
150         Call WriteMacroTrabajoToggle(UserIndex, True)

        End If
    
        
        Exit Sub

DoLingotes_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoLingotes", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ModFundicion", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ModAlquimia", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ModSastre", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ModCarpinteria", Erl)
        Resume Next
        
End Function

Function ModHerreriA(ByVal clase As eClass) As Single
        
        On Error GoTo ModHerreriA_Err
        

100     Select Case clase

            Case eClass.Trabajador
102             ModHerreriA = 1

104         Case Else
106             ModHerreriA = 3

        End Select

        
        Exit Function

ModHerreriA_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ModHerreriA", Erl)
        Resume Next
        
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
        
        On Error GoTo DoAdminInvisible_Err
        
    
100     If UserList(UserIndex).flags.AdminInvisible = 0 Then
                
102         UserList(UserIndex).flags.AdminInvisible = 1
104         UserList(UserIndex).flags.invisible = 1
106         UserList(UserIndex).flags.Oculto = 1
108         UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.Body
110         UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
112         UserList(UserIndex).Char.Body = 0
114         UserList(UserIndex).Char.Head = 0
        
        Else
        
116         UserList(UserIndex).flags.AdminInvisible = 0
118         UserList(UserIndex).flags.invisible = 0
120         UserList(UserIndex).flags.Oculto = 0
122         UserList(UserIndex).Counters.TiempoOculto = 0
124         UserList(UserIndex).Char.Body = UserList(UserIndex).flags.OldBody
126         UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
        End If
    
        'vuelve a ser visible por la fuerza
128     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
130     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))

        
        Exit Sub

DoAdminInvisible_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoAdminInvisible", Erl)
        Resume Next
        
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo TratarDeHacerFogata_Err
        

        Dim Suerte    As Byte

        Dim exito     As Byte

        Dim obj       As obj

        Dim posMadera As WorldPos

100     If Not LegalPos(Map, x, Y) Then Exit Sub

102     With posMadera
104         .Map = Map
106         .x = x
108         .Y = Y

        End With

110     If MapData(Map, x, Y).ObjInfo.ObjIndex <> 58 Then
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

122     If MapData(Map, x, Y).ObjInfo.Amount < 3 Then
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
144         obj.Amount = MapData(Map, x, Y).ObjInfo.Amount \ 3
    
146         Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.Amount & " ramitas.", FontTypeNames.FONTTYPE_INFO)
    
148         Call MakeObj(obj, Map, x, Y)
    
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.TratarDeHacerFogata", Erl)
        Resume Next
        
End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal RedDePesca As Boolean = False, Optional ByVal ObjetoDorado As Boolean = False)

    On Error GoTo Errhandler

    Dim Suerte       As Integer
    Dim res          As Integer
    Dim RestaStamina As Byte

    RestaStamina = IIf(RedDePesca, 2, 1)
    
    With UserList(UserIndex)
    
        If .Stats.MinSta > RestaStamina Then
            Call QuitarSta(UserIndex, RestaStamina)
        
        Else
            
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para pescar.", FontTypeNames.FONTTYPE_INFO)
            
            Call WriteMacroTrabajoToggle(UserIndex, False)
            
            Exit Sub

        End If

        Dim Skill As Integer

        Skill = .Stats.UserSkills(eSkill.Pescar)
        
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))

        If res < 6 Then

            Dim nPos  As WorldPos
            Dim MiObj As obj
        
            MiObj.Amount = IIf(RedDePesca, RandomNumber(2, 5), IIf(ObjetoDorado, RandomNumber(1, 3), 1))
            MiObj.ObjIndex = ObtenerPezRandom(ObjData(.Invent.HerramientaEqpObjIndex).Power)
        
            If MiObj.ObjIndex = 0 Then Exit Sub

            If ObjData(.Invent.HerramientaEqpObjIndex).donador = 1 Then
                MiObj.Amount = MiObj.Amount * 2
                MiObj.Amount = MiObj.Amount * RecoleccionMult
            Else
                MiObj.Amount = MiObj.Amount * RecoleccionMult
            End If
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)
            End If

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(.Pos.x, .Pos.Y, MiObj.Amount, 5))
        
            ' Al pescar también podés sacar cosas raras (se setean desde RecursosEspeciales.dat)
            Dim i As Integer

            ' Por cada drop posible
            For i = 1 To UBound(EspecialesPesca)
                ' Tiramos al azar entre 1 y la probabilidad
                res = RandomNumber(1, IIf(RedDePesca, EspecialesPesca(i).data * 2, EspecialesPesca(i).data)) ' Red de pesca chance x2 (revisar)
            
                ' Si tiene suerte y le pega
                If res = 1 Then
                    MiObj.ObjIndex = EspecialesPesca(i).ObjIndex
                    MiObj.Amount = 1 ' Solo un item por vez
                
                    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                    
                    ' Le mandamos un mensaje
                    Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(EspecialesPesca(i).ObjIndex).name & "!", FontTypeNames.FONTTYPE_INFO)

                    ' TODO: Sonido ?
                    'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                End If

            Next

        End If
    
        Call SubirSkill(UserIndex, eSkill.Pescar)
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    
        'Ladder 06/07/14 Activamos el macro de trabajo
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " - " & Err.description)

End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo DoRobar_Err
        

100     If Not MapInfo(UserList(victimaindex).Pos.Map).Seguro = 0 Then Exit Sub

102     If UserList(LadrOnIndex).flags.Seguro Then
104         Call WriteConsoleMsg(LadrOnIndex, "Debes quitar el seguro para robar", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

106     If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then
108         Call WriteConsoleMsg(LadrOnIndex, "Para robar debes salir de la armada real.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

110     If TriggerZonaPelea(LadrOnIndex, victimaindex) <> TRIGGER6_AUSENTE Then Exit Sub

112     If UserList(LadrOnIndex).GuildIndex > 0 Then
114         If UserList(LadrOnIndex).flags.SeguroClan Then
116             If UserList(LadrOnIndex).GuildIndex = UserList(victimaindex).GuildIndex Then
118                 Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If

        End If

120     If UserList(LadrOnIndex).Grupo.EnGrupo > 0 Then
122         If UserList(LadrOnIndex).GuildIndex = UserList(victimaindex).GuildIndex Then
124             Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

        End If

126     If UserList(LadrOnIndex).Grupo.EnGrupo = True Then

            Dim i As Byte

128         For i = 1 To UserList(UserList(LadrOnIndex).Grupo.Lider).Grupo.CantidadMiembros

130             If UserList(UserList(LadrOnIndex).Grupo.Lider).Grupo.Miembros(i) = victimaindex Then
132                 Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

134         Next i

        End If

136     Call QuitarSta(LadrOnIndex, 15)

138     If UserList(victimaindex).flags.Privilegios And PlayerType.user Then

            Dim Suerte     As Integer

            Dim res        As Integer

            Dim Porcentaje As Byte
    
140         If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
142             Suerte = 35
144             Porcentaje = 1
146         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
148             Suerte = 30
150             Porcentaje = 1
152         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
154             Suerte = 28
156             Porcentaje = 2
158         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
160             Suerte = 24
162             Porcentaje = 3
164         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
166             Suerte = 22
168             Porcentaje = 4
170         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
172             Suerte = 20
174             Porcentaje = 5
176         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
178             Suerte = 18
180             Porcentaje = 6
182         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
184             Suerte = 15
186             Porcentaje = 7
188         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
190             Suerte = 10
192             Porcentaje = 8
194         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
196             Suerte = 7
198             Porcentaje = 9
200         ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
202             Suerte = 5
204             Porcentaje = 10

            End If

206         res = RandomNumber(1, Suerte)
        
208         If res < 4 Then 'Exito robo
    
                ' TODO: Clase ladrón
                'If UserList(LadrOnIndex).clase = eClass.Trabajador Then
210             If False Then
           
212                 If (RandomNumber(1, 50) < 25) Then
214                     If TieneObjetosRobables(victimaindex) Then
216                         Call RobarObjeto(LadrOnIndex, victimaindex)
                        Else
218                         Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else 'Roba oro

220                     If UserList(victimaindex).Stats.GLD > 0 Then

                            Dim n As Long
                    
                            'porcentaje
222                         n = UserList(victimaindex).Stats.GLD / 100 * Porcentaje
                    
                            ' N = RandomNumber(100, 1000)
224                         If n > UserList(victimaindex).Stats.GLD Then n = UserList(victimaindex).Stats.GLD
226                         UserList(victimaindex).Stats.GLD = UserList(victimaindex).Stats.GLD - n
                        
228                         UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n

230                         If UserList(LadrOnIndex).Stats.GLD > MAXORO Then UserList(LadrOnIndex).Stats.GLD = MAXORO
                        
232                         Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(victimaindex).name, FontTypeNames.FONTTYPE_INFO)
                        Else
234                         Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                Else

236                 If UserList(victimaindex).Stats.GLD > 0 Then

238                     n = UserList(victimaindex).Stats.GLD / 100 * 0.5
                    
240                     If n > UserList(victimaindex).Stats.GLD Then n = UserList(victimaindex).Stats.GLD
242                     UserList(victimaindex).Stats.GLD = UserList(victimaindex).Stats.GLD - n
                        
244                     UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n

246                     If UserList(LadrOnIndex).Stats.GLD > MAXORO Then UserList(LadrOnIndex).Stats.GLD = MAXORO
                        
248                     Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(victimaindex).name, FontTypeNames.FONTTYPE_INFO)
                    Else
250                     Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
        
            End If
    
        Else
252         Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
254         Call WriteConsoleMsg(victimaindex, "¡" & UserList(LadrOnIndex).name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
256         Call WriteConsoleMsg(victimaindex, "¡" & UserList(LadrOnIndex).name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
        

        End If

258     If Status(LadrOnIndex) = 1 Then
260         Call VolverCriminal(LadrOnIndex)

        End If
    
262     If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

264     Call SubirSkill(LadrOnIndex, Robar)
266     Call WriteUpdateGold(LadrOnIndex)
268     Call WriteUpdateGold(victimaindex)

        
        Exit Sub

DoRobar_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoRobar", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ObjEsRobable", Erl)
        Resume Next
        
End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo RobarObjeto_Err
        

        'Call LogTarea("Sub RobarObjeto")
        Dim flag As Boolean

        Dim i    As Integer

100     flag = False

102     If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
104         i = 1

106         Do While Not flag And i <= MAX_INVENTORY_SLOTS

                'Hay objeto en este slot?
108             If UserList(victimaindex).Invent.Object(i).ObjIndex > 0 Then
110                 If ObjEsRobable(victimaindex, i) Then
112                     If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

114             If Not flag Then i = i + 1
            Loop
        Else
116         i = 20

118         Do While Not flag And i > 0

                'Hay objeto en este slot?
120             If UserList(victimaindex).Invent.Object(i).ObjIndex > 0 Then
122                 If ObjEsRobable(victimaindex, i) Then
124                     If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

126             If Not flag Then i = i - 1
            Loop

        End If

128     If flag Then

            Dim MiObj As obj

            Dim num   As Byte

            'Cantidad al azar
130         num = RandomNumber(1, 5)
                
132         If num > UserList(victimaindex).Invent.Object(i).Amount Then
134             num = UserList(victimaindex).Invent.Object(i).Amount

            End If
                
136         MiObj.Amount = num
138         MiObj.ObjIndex = UserList(victimaindex).Invent.Object(i).ObjIndex
    
140         UserList(victimaindex).Invent.Object(i).Amount = UserList(victimaindex).Invent.Object(i).Amount - num
                
142         If UserList(victimaindex).Invent.Object(i).Amount <= 0 Then
144             Call QuitarUserInvItem(victimaindex, CByte(i), 1)

            End If
            
146         Call UpdateUserInv(False, victimaindex, CByte(i))
                
148         If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
150             Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

            End If
    
152         Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)

        Else
154         Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If exiting, cancel de quien es robado
156     Call CancelExit(victimaindex)

        
        Exit Sub

RobarObjeto_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.RobarObjeto", Erl)
        Resume Next
        
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

            Case eClass.Assasin '35
104             Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 4.25)
        
106         Case eClass.Cleric, eClass.Paladin ' 15
108             Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
        
110         Case eClass.Bard, eClass.Druid '13
112             Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
        
114         Case Else '8
116             Suerte = Int(0.0361 * Skill + 4.39)

        End Select
    
118     If RandomNumber(0, 70) < Suerte Then
120         If VictimUserIndex <> 0 Then
                daño = daño * 1.5
            
128             UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño

130             Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageEfectOverHead("¡" & daño & "!", UserList(VictimUserIndex).Char.CharIndex, &HFFFF00))

132             If UserList(UserIndex).ChatCombate = 1 Then
                    'Call WriteEfectOverHead(UserIndex, daño, UserList(UserIndex).Char.CharIndex) 'LADDER 21.11.08
134                 Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

                End If

136             If UserList(VictimUserIndex).ChatCombate = 1 Then
138                 Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
            
            Else
                daño = daño * 2

140             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño

142             If UserList(UserIndex).ChatCombate = 1 Then
                    'Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 1.5), FontTypeNames.FONTTYPE_FIGHT)
144                 Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, daño)

                End If
            
146             Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageEfectOverHead("¡" & daño & "!", Npclist(VictimNpcIndex).Char.CharIndex, &HFFFF00))

                '[Alejo]
148             Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)

            End If

        Else

150         If UserList(UserIndex).ChatCombate = 1 Then
152             Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If
    
154     Call SubirSkill(UserIndex, Apuñalar)

        
        Exit Sub

DoApuñalar_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoApuñalar", Erl)
        Resume Next
        
End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
        
        On Error GoTo DoGolpeCritico_Err
        

        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 28/01/2007
        '***************************************************
        Dim Suerte As Integer

        Dim Skill  As Integer

        'If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
100     If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
102     If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).name <> "Espada Vikinga" Then Exit Sub

104     Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

106     Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

108     If RandomNumber(0, 100) < Suerte Then
110         daño = Int(daño * 0.5)

112         If VictimUserIndex <> 0 Then
114             UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
116             Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
118             Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Else
120             Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
122             Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
                '[Alejo]
124             Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)

            End If

        End If

        
        Exit Sub

DoGolpeCritico_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoGolpeCritico", Erl)
        Resume Next
        
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarSta_Err
        
100     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

102     If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
104     If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub
106     Call WriteUpdateSta(UserIndex)

        
        Exit Sub

QuitarSta_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.QuitarSta", Erl)
        Resume Next
        
End Sub

Public Sub DoRaices(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte)

    On Error GoTo Errhandler

    Dim Suerte As Integer
    Dim res    As Integer
    
    With UserList(UserIndex)
    
        If .Stats.MinSta > 2 Then
            Call QuitarSta(UserIndex, 2)
        
        Else
            
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para obtener raices.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
    
        End If
    
        Dim Skill As Integer
            Skill = .Stats.UserSkills(eSkill.Alquimia)
        
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        res = RandomNumber(1, Suerte)
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
    
        Rem Ladder 06/08/14 Subo un poco la probabilidad de sacar raices... porque era muy lento
        If res < 7 Then
    
            Dim nPos  As WorldPos
            Dim MiObj As obj
        
            'If .clase = eClass.Druid Then
            'MiObj.Amount = RandomNumber(6, 8)
            ' Else
            MiObj.Amount = RandomNumber(5, 7)
            ' End If
       
            If ObjData(.Invent.HerramientaEqpObjIndex).donador = 1 Then
                MiObj.Amount = MiObj.Amount * 2
            End If
       
            MiObj.Amount = MiObj.Amount * RecoleccionMult
            MiObj.ObjIndex = Raices
        
            MapData(.Pos.Map, x, Y).ObjInfo.Amount = MapData(.Pos.Map, x, Y).ObjInfo.Amount - MiObj.Amount
    
            If MapData(.Pos.Map, x, Y).ObjInfo.Amount < 0 Then
                MapData(.Pos.Map, x, Y).ObjInfo.Amount = 0
    
                ' VidaUtil.Item_ListAdd .Pos.Map, X, Y
            End If
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
            
                Call TirarItemAlPiso(.Pos, MiObj)
            
            End If
        
            'Call WriteConsoleMsg(UserIndex, "¡Has conseguido algunas raices!", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(.Pos.x, .Pos.Y, MiObj.Amount, 5))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(60, .Pos.x, .Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(61, .Pos.x, .Pos.Y))
    
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 8 Then
                Call WriteConsoleMsg(UserIndex, "¡No has obtenido raices!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 8
    
            End If
        
            '[/CDT]
        End If
    
        Call SubirSkill(UserIndex, eSkill.Alquimia)
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en DoRaices")

End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

    On Error GoTo Errhandler

    Dim Suerte As Integer
    Dim res    As Integer
    
    With UserList(UserIndex)
    
            'EsfuerzoTalarLeñador = 1
        If .Stats.MinSta > 2 Then
            Call QuitarSta(UserIndex, 2)
        
        Else
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para talar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
    
        End If
    
        Dim Skill As Integer
    
        Skill = .Stats.UserSkills(eSkill.Talar)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
        res = RandomNumber(1, Suerte)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        
        If res < 6 Then
    
            Dim nPos  As WorldPos
    
            Dim MiObj As obj
            
            If .flags.TargetObj = 0 Then Exit Sub
            
            Call ActualizarRecurso(.Pos.Map, x, Y)
            MapData(.Pos.Map, x, Y).ObjInfo.data = timeGetTime ' Ultimo uso
    
            MiObj.Amount = IIf(ObjetoDorado, RandomNumber(1, 5), 1)
            
            If ObjData(.Invent.HerramientaEqpObjIndex).donador = 1 Then
                MiObj.Amount = MiObj.Amount * 2
    
            End If
            
            MiObj.Amount = MiObj.Amount * RecoleccionMult
            MiObj.ObjIndex = Leña
            
            If MiObj.Amount > MapData(.Pos.Map, x, Y).ObjInfo.Amount Then
                MiObj.Amount = MapData(.Pos.Map, x, Y).ObjInfo.Amount
    
            End If
            
            MapData(.Pos.Map, x, Y).ObjInfo.Amount = MapData(.Pos.Map, x, Y).ObjInfo.Amount - MiObj.Amount
            
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                
                Call TirarItemAlPiso(.Pos, MiObj)
                
            End If
    
            'If Not .flags.UltimoMensaje = 5 Then
            ' Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
            '        .flags.UltimoMensaje = 5
            ' End If
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(.Pos.x, .Pos.Y, MiObj.Amount, 5))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.x, .Pos.Y))
            
            ' Al talar también podés dropear cosas raras (se setean desde RecursosEspeciales.dat)
            Dim i As Integer
    
            ' Por cada drop posible
            For i = 1 To UBound(EspecialesTala)
                ' Tiramos al azar entre 1 y la probabilidad
                res = RandomNumber(1, EspecialesTala(i).data)
                
                ' Si tiene suerte y le pega
                If res = 1 Then
                    MiObj.ObjIndex = EspecialesTala(i).ObjIndex
                    MiObj.Amount = 1 ' Solo un item por vez
                    
                    'If Not MeterItemEnInventario(Userindex, MiObj) Then _
                    'Call TirarItemAlPiso(.Pos, MiObj)
    
                    ' Tiro siempre el item al piso, me parece más rolero, como que cae del árbol :P
                    Call TirarItemAlPiso(.Pos, MiObj)
    
                    ' Oculto el mensaje porque el item cae al piso
                    'Call WriteConsoleMsg(Userindex, "¡Has conseguido " & ObjData(EspecialesTala(i).ObjIndex).Name & "!", FontTypeNames.FONTTYPE_INFO)
                    ' TODO: Sonido ?
                    'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                End If
    
            Next
        
        Else
            '[CDT 17-02-2004]
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(64, .Pos.x, .Pos.Y))
    
            If Not .flags.UltimoMensaje = 8 Then
                Call WriteConsoleMsg(UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 8
    
            End If
    
            '[/CDT]
        End If
        
        Call SubirSkill(UserIndex, eSkill.Talar)
        
        .Counters.Trabajando = .Counters.Trabajando + 1
    
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    
    End With

    Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

    On Error GoTo Errhandler

    Dim Suerte As Integer
    Dim res    As Integer
    Dim metal  As Integer

    With UserList(UserIndex)
    
        'Por Ladder 06/07/2014 Cuando la estamina llega a 0 , el macro se desactiva
        If .Stats.MinSta > 2 Then
            Call QuitarSta(UserIndex, 2)
        Else
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
    
        End If
    
        'Por Ladder 06/07/2014
    
        Dim Skill As Integer
    
        Skill = .Stats.UserSkills(eSkill.Mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
        
        res = RandomNumber(1, Suerte)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
        
        If res <= 5 Then
    
            Dim MiObj As obj
            Dim nPos  As WorldPos
            
            If .flags.TargetObj = 0 Then Exit Sub
            
            Call ActualizarRecurso(.Pos.Map, x, Y)
            MapData(.Pos.Map, x, Y).ObjInfo.data = timeGetTime ' Ultimo uso
            
            MiObj.ObjIndex = ObjData(.flags.TargetObj).MineralIndex
            
            MiObj.Amount = IIf(ObjetoDorado, RandomNumber(1, 6), 1)
    
            If ObjData(.Invent.HerramientaEqpObjIndex).donador = 1 Then
                MiObj.Amount = MiObj.Amount * 2
    
            End If
            
            MiObj.Amount = MiObj.Amount * RecoleccionMult
            
            If MiObj.Amount > MapData(.Pos.Map, x, Y).ObjInfo.Amount Then
                MiObj.Amount = MapData(.Pos.Map, x, Y).ObjInfo.Amount
    
            End If
            
            MapData(.Pos.Map, x, Y).ObjInfo.Amount = MapData(.Pos.Map, x, Y).ObjInfo.Amount - MiObj.Amount
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
            
            Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
            
            ' Al minar también puede dropear una gema
            Dim i As Integer
    
            ' Por cada drop posible
            For i = 1 To ObjData(.flags.TargetObj).CantItem
                ' Tiramos al azar entre 1 y la probabilidad
                res = RandomNumber(1, ObjData(.flags.TargetObj).Item(i).Amount)
                
                ' Si tiene suerte y le pega
                If res = 1 Then
                    ' Se lo metemos al inventario (o lo tiramos al piso)
                    MiObj.ObjIndex = ObjData(.flags.TargetObj).Item(i).ObjIndex
                    MiObj.Amount = 1 ' Solo una gema por vez
                    
                    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                        
                    ' Le mandamos un mensaje
                    Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(ObjData(.flags.TargetObj).Item(i).ObjIndex).name & "!", FontTypeNames.FONTTYPE_INFO)
                    ' TODO: Sonido de drop de gema :P
                    'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(15, .Pos.x, .Pos.Y))
                        
                    ' Como máximo dropea una gema
                    'Exit For ' Lo saco a pedido de Haracin
                End If
    
            Next
            
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(62, .Pos.x, .Pos.Y))
    
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                
                Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                
                .flags.UltimoMensaje = 9
    
            End If
    
            '[/CDT]
        End If
        
        Call SubirSkill(UserIndex, eSkill.Mineria)
        
        .Counters.Trabajando = .Counters.Trabajando + 1
        
        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
        
        On Error GoTo DoMeditar_Err
        

        Dim Suerte       As Integer

        Dim res          As Integer

        Dim cant         As Integer

        Dim MeditarSkill As Byte

100     With UserList(UserIndex)
            '.Counters.IdleCount = 0

102         If .Stats.MinMAN >= .Stats.MaxMAN Then Exit Sub
    
104         MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
106         If MeditarSkill <= 10 Then
108             Suerte = 35
110         ElseIf MeditarSkill <= 20 Then
112             Suerte = 30
114         ElseIf MeditarSkill <= 30 Then
116             Suerte = 28
118         ElseIf MeditarSkill <= 40 Then
120             Suerte = 24
122         ElseIf MeditarSkill <= 50 Then
124             Suerte = 22
126         ElseIf MeditarSkill <= 60 Then
128             Suerte = 20
130         ElseIf MeditarSkill <= 70 Then
132             Suerte = 18
134         ElseIf MeditarSkill <= 80 Then
136             Suerte = 15
138         ElseIf MeditarSkill <= 90 Then
140             Suerte = 10
142         ElseIf MeditarSkill < 100 Then
144             Suerte = 7
            Else
146             Suerte = 5
            End If
    
148         If .flags.RegeneracionMana = 1 Then
150             Suerte = 10
            End If
        
152         res = RandomNumber(1, Suerte)
    
154         If res = 1 Then
156             cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)

158             If cant <= 0 Then cant = 1
160             .Stats.MinMAN = .Stats.MinMAN + cant

162             If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
                '  If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
                '     Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
                '     UserList(UserIndex).flags.UltimoMensaje = 22
                '  End If
            
164             Call WriteUpdateMana(UserIndex)
166             Call SubirSkill(UserIndex, Meditar)

            End If

        End With

        
        Exit Sub

DoMeditar_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoMeditar", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.Desarmar", Erl)
        Resume Next
        
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
114         If (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger > 10) Then
116             Call WriteConsoleMsg(UserIndex, "No podés montar aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

118     If UserList(UserIndex).flags.Meditando Then
120         UserList(UserIndex).flags.Meditando = False
122         Call WriteLocaleMsg(UserIndex, "123", FontTypeNames.FONTTYPE_INFO)
124         UserList(UserIndex).Char.FX = 0
126         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
        End If

128     If UserList(UserIndex).flags.Montado = 1 Then
130         If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
132             If ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
134                 UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica
136                 Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MonturaSlot)

                End If

            End If

        End If

138     UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
140     UserList(UserIndex).Invent.MonturaSlot = slot

142     If UserList(UserIndex).flags.Montado = 0 Then

144         If ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
146             UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + Montura.ResistenciaMagica

            End If
            
148         UserList(UserIndex).Char.Body = Montura.Ropaje

            'UserList(UserIndex).Char.body = Montura.Ropaje
154         UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
156         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
158         UserList(UserIndex).Char.WeaponAnim = NingunArma
160         UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).Char.CascoAnim
162         UserList(UserIndex).flags.Montado = 1
164         UserList(UserIndex).Char.speeding = VelocidadMontura
        Else
166         UserList(UserIndex).flags.Montado = 0
168         UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
170         UserList(UserIndex).Char.speeding = VelocidadNormal

172         If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

            Else
180             Call DarCuerpoDesnudo(UserIndex)

            End If
            
182         If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

184         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

186         If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

        End If

188     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

190     Call UpdateUserInv(False, UserIndex, slot)
192     Call WriteEquiteToggle(UserIndex)
194     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

        
        Exit Sub

DoMontar_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.DoMontar", Erl)
        Resume Next
        
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
104             Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 4.25)
        
106             If VictimNpcIndex = 0 Then
108                 If UserList(VictimUserIndex).Char.heading = UserList(UserIndex).Char.heading Then
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

134     If RandomNumber(0, 70) < Suerte Then
136         If VictimUserIndex <> 0 Then
                ApuñalarFunction = daño * 1.5
            Else
                ApuñalarFunction = daño * 2
            End If
        End If
        
        Exit Function

ApuñalarFunction_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ApuñalarFunction", Erl)
        Resume Next
        
End Function

Public Sub ActualizarRecurso(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)
        
        On Error GoTo ActualizarRecurso_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = MapData(Map, x, Y).ObjInfo.ObjIndex

        Dim TiempoActual As Long

102     TiempoActual = timeGetTime

        ' Data = Ultimo uso
104     If (TiempoActual - MapData(Map, x, Y).ObjInfo.data) * 0.001 > ObjData(ObjIndex).TiempoRegenerar Then
106         MapData(Map, x, Y).ObjInfo.Amount = ObjData(ObjIndex).VidaUtil
108         MapData(Map, x, Y).ObjInfo.data = &H7FFFFFFF   ' Ultimo uso = Max Long

        End If

        
        Exit Sub

ActualizarRecurso_Err:
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ActualizarRecurso", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "Trabajo.ObtenerPezRandom", Erl)
        Resume Next
        
End Function

Function ModDomar(ByVal clase As eClass) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case clase

        Case eClass.Druid
            ModDomar = 6

        Case eClass.Hunter
            ModDomar = 6

        Case eClass.Cleric
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Nacho (Integer)
    'Last Modification: 01/05/2010
    '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
    '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
    '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
    '***************************************************

    On Error GoTo Errhandler

    Dim puntosDomar      As Integer

    Dim puntosRequeridos As Integer

    Dim CanStay          As Boolean

    Dim petType          As Integer

    Dim NroPets          As Integer
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    With UserList(UserIndex)

        If .NroMascotas < MAXMASCOTAS Then

            If Npclist(NpcIndex).MaestroNPC > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))

            ' 20% de bonificacion
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8

            ' 11% de bonificacion
            ElseIf .Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
                puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.89

            Else
                puntosRequeridos = Npclist(NpcIndex).flags.Domable
            End If

            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then

                Dim Index As Integer

                .NroMascotas = .NroMascotas + 1
                Index = FreeMascotaIndex(UserIndex)
                .MascotasIndex(Index) = NpcIndex
                .MascotasType(Index) = Npclist(NpcIndex).Numero

                Npclist(NpcIndex).MaestroUser = UserIndex

                Call FollowAmo(NpcIndex)
                Call ReSpawnNpc(Npclist(NpcIndex))

                Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)

                ' Es zona segura?
                If MapInfo(.Pos.Map).Seguro Then
                    petType = Npclist(NpcIndex).Numero
                    NroPets = .NroMascotas

                    Call QuitarNPC(NpcIndex)

                    .MascotasType(Index) = petType
                    .NroMascotas = NroPets

                    Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)
                End If

            Else

                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5
                End If

            End If

            Call SubirSkill(UserIndex, eSkill.Domar)

        Else
            Call WriteConsoleMsg(UserIndex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long

    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1

        End If

    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Private Function ModFundirMineral(ByVal clase As eClass) As Integer
    
    If clase = eClass.Trabajador Then
        ModFundirMineral = 1
    Else
        ModFundirMineral = 3
    End If
    
End Function
