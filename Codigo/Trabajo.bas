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

    UserList(UserIndex).Counters.TiempoOculto = UserList(UserIndex).Counters.TiempoOculto - 1

    If UserList(UserIndex).Counters.TiempoOculto <= 0 Then
    
        ' UserList(UserIndex).Counters.TiempoOculto = IntervaloOculto
    
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).flags.Oculto = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
        Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

    End If

    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")

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

    Dim ModNave As Long

    If UserList(UserIndex).flags.Nadando = 0 Then
    
        If UserList(UserIndex).flags.Muerto = 0 Then
            '(Nacho)
    
            UserList(UserIndex).Char.Body = 694
            'If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.body = iGalera
            'If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.body = iGaleon
        Else
            UserList(UserIndex).Char.Body = iFragataFantasmal

        End If
    
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
        UserList(UserIndex).flags.Nadando = 1
    
    Else
    
        UserList(UserIndex).flags.Nadando = 0
    
        If UserList(UserIndex).flags.Muerto = 0 Then
            UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            
                If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
                    UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).RopajeBajo
                Else
                    UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

                End If
            
            Else
                Call DarCuerpoDesnudo(UserIndex)

            End If
        
            If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
        Else
            UserList(UserIndex).Char.Body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco

        End If

    End If

    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    'Call WriteNadarToggle(UserIndex)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)

    Dim ModNave As Long

    If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
    UserList(UserIndex).Invent.BarcoSlot = slot

    If UserList(UserIndex).flags.Montado > 0 Then
        Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

    End If

    If UserList(UserIndex).flags.Navegando = 0 Then

        If Barco.Ropaje = iTraje Then
            Call WriteNadarToggle(UserIndex, True)
        
        Else
            Call WriteNadarToggle(UserIndex, False)
        
        End If
    
        If Barco.Ropaje <> iTraje Then
            UserList(UserIndex).Char.Head = 0
            UserList(UserIndex).Char.CascoAnim = NingunCasco

        End If
    
        If UserList(UserIndex).flags.Muerto = 0 Then

            '(Nacho)
            If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaCiuda
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraCiuda
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonCiuda
            ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then

                If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaPk
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraPk
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonPk
            Else

                If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
                If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarca
                If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGalera
                If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleon

            End If

        Else

            If Barco.Ropaje = iTraje Then
                UserList(UserIndex).Char.Body = iRopaBuceoMuerto
            Else
                UserList(UserIndex).Char.Body = iFragataFantasmal

            End If

            UserList(UserIndex).Char.Head = iCabezaMuerto

        End If
    
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        'UserList(UserIndex).Char.CascoAnim = NingunCasco
        UserList(UserIndex).flags.Navegando = 1
    
        UserList(UserIndex).Char.speeding = Barco.Velocidad
    
    Else

        If Barco.Ropaje = iTraje Then
            Call WriteNadarToggle(UserIndex, False)
        Else
            Call WriteNadarToggle(UserIndex, False)

        End If

        UserList(UserIndex).Char.speeding = VelocidadNormal
    
        UserList(UserIndex).flags.Navegando = 0
    
        If UserList(UserIndex).flags.Muerto = 0 Then
            UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
            
                If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
                    UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).RopajeBajo
                Else
                    UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

                End If
            
            Else
                Call DarCuerpoDesnudo(UserIndex)

            End If
        
            If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

            If UserList(UserIndex).Invent.NudilloObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.NudilloObjIndex).WeaponAnim

            If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).WeaponAnim

            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
        Else
            UserList(UserIndex).Char.Body = iCuerpoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            UserList(UserIndex).Char.CascoAnim = NingunCasco

        End If

    End If

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

    'Call WriteVelocidadToggle(UserIndex)
    
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call WriteNavigateToggle(UserIndex)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

End Sub

Public Sub DoReNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal slot As Integer)

    Dim ModNave As Long

    If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < Barco.MinSkill Then
        Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
    UserList(UserIndex).Invent.BarcoSlot = slot

    If UserList(UserIndex).flags.Montado > 0 Then
        Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

    End If

    If Barco.Ropaje = iTraje Then
        Call WriteNadarToggle(UserIndex, True)
    Else
        Call WriteNadarToggle(UserIndex, False)

    End If
    
    If Barco.Ropaje <> iTraje Then
        UserList(UserIndex).Char.Head = 0
    Else
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head

    End If
    
    If UserList(UserIndex).flags.Muerto = 0 Then

        '(Nacho)
        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
            If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
            If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaCiuda
            If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraCiuda
            If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonCiuda
        ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then

            If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
            If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaPk
            If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraPk
            If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonPk
        Else

            If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
            If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarca
            If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGalera
            If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleon

        End If

    Else
        UserList(UserIndex).Char.Body = iFragataFantasmal

    End If
    
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).flags.Navegando = 1
    
    UserList(UserIndex).Char.speeding = Barco.Velocidad

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

    '
    'Call WriteVelocidadToggle(UserIndex)
    
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.BARCA_SOUND, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
        If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
            ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) Then
            
            Call DoLingotes(UserIndex)
        
        Else
            Call WriteConsoleMsg(UserIndex, "No tenes conocimientos de mineria suficientes para trabajar este mineral. Necesitas " & ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill & " puntos en mineria.", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
    'Call LogTarea("Sub TieneObjetos")

    Dim i     As Long

    Dim Total As Long

    For i = 1 To UserList(UserIndex).CurrentInventorySlots

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount

        End If

    Next i

    If cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
    'Call LogTarea("Sub QuitarObjetos")

    Dim i As Long

    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        Debug.Print i

        If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Debug.Print UserList(UserIndex).name
        
            If UserList(UserIndex).Invent.Object(i).Equipped Then
                Call Desequipar(UserIndex, i)

            End If
        
            UserList(UserIndex).Invent.Object(i).Amount = UserList(UserIndex).Invent.Object(i).Amount - cant

            If (UserList(UserIndex).Invent.Object(i).Amount <= 0) Then
                cant = Abs(UserList(UserIndex).Invent.Object(i).Amount)
                UserList(UserIndex).Invent.Object(i).Amount = 0
                UserList(UserIndex).Invent.Object(i).ObjIndex = 0
            Else
                cant = 0

            End If
        
            Call UpdateUserInv(False, UserIndex, i)
        
            If (cant = 0) Then
                QuitarObjetos = True
                Exit Function

            End If

        End If

    Next i

End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)

End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex)

End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)

End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex)
    If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
    If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex)

End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
        If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
            CarpinteroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If
    
    CarpinteroTieneMateriales = True

End Function

Function AlquimistaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Raices > 0 Then
        If Not TieneObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes raices.", FontTypeNames.FONTTYPE_INFO)
            AlquimistaTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If
    
    AlquimistaTieneMateriales = True

End Function

Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).PielLobo > 0 Then
        If Not TieneObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de lobo.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If
    
    If ObjData(ItemIndex).PielOsoPardo > 0 Then
        If Not TieneObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso pardo.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If
    
    If ObjData(ItemIndex).PielOsoPolaR > 0 Then
        If Not TieneObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso polar.", FontTypeNames.FONTTYPE_INFO)
            SastreTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If
    
    SastreTieneMateriales = True

End Function

Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean

    If ObjData(ItemIndex).LingH > 0 Then
        If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If

    If ObjData(ItemIndex).LingP > 0 Then
        If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If

    If ObjData(ItemIndex).LingO > 0 Then
        If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
            HerreroTieneMateriales = False
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Function

        End If

    End If

    HerreroTieneMateriales = True

End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
        Call HerreroQuitarMateriales(UserIndex, ItemIndex)
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        Call WriteUpdateSta(UserIndex)
        ' AGREGAR FX
    
        If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
            ' Call WriteConsoleMsg(UserIndex, "Has construido el arma!", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
            ' Call WriteConsoleMsg(UserIndex, "Has construido el escudo!", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
            ' Call WriteConsoleMsg(UserIndex, "Has construido el casco!", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
            'Call WriteConsoleMsg(UserIndex, "Has construido la armadura!", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))

        End If

        Dim MiObj As obj

        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

        End If
    
        'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
        ' If ObjData(MiObj.ObjIndex).Log = 1 Then
        '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
        'End If
    
        Call SubirSkill(UserIndex, eSkill.Herreria)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    End If

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long

    For i = 1 To UBound(ObjCarpintero)

        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function

        End If

    Next i

    PuedeConstruirCarpintero = False

End Function

Public Function PuedeConstruirAlquimista(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long

    For i = 1 To UBound(ObjAlquimista)

        If ObjAlquimista(i) = ItemIndex Then
            PuedeConstruirAlquimista = True
            Exit Function

        End If

    Next i

    PuedeConstruirAlquimista = False

End Function

Public Function PuedeConstruirSastre(ByVal ItemIndex As Integer) As Boolean

    Dim i As Long

    For i = 1 To UBound(ObjSastre)

        If ObjSastre(i) = ItemIndex Then
            PuedeConstruirSastre = True
            Exit Function

        End If

    Next i

    PuedeConstruirSastre = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If CarpinteroTieneMateriales(UserIndex, ItemIndex) And _
        UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria And _
        PuedeConstruirCarpintero(ItemIndex) And _
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then
    
        If UserList(UserIndex).Stats.MinSta > 2 Then
            Call QuitarSta(UserIndex, 2)
        
        Else
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub

        End If
    
        Call CarpinteroQuitarMateriales(UserIndex, ItemIndex)
        'Call WriteConsoleMsg(UserIndex, "Has construido un objeto!", FontTypeNames.FONTTYPE_INFO)
        'Call WriteOroOverHead(UserIndex, 1, UserList(UserIndex).Char.CharIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
    
        Dim MiObj As obj

        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

        End If
    
        'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
        ' If ObjData(MiObj.ObjIndex).Log = 1 Then
        '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
        ' End If
    
        Call SubirSkill(UserIndex, eSkill.Carpinteria)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    End If

End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    Rem Debug.Print UserList(UserIndex).Invent.HerramientaEqpObjIndex

    If Not UserList(UserIndex).Stats.MinSta > 0 Then
        Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If AlquimistaTieneMateriales(UserIndex, ItemIndex) And _
        UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) >= ObjData(ItemIndex).SkPociones And _
        PuedeConstruirAlquimista(ItemIndex) And UserList(UserIndex).Invent.HerramientaEqpObjIndex = OLLA_ALQUIMIA Then
        
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 25
        Call WriteUpdateSta(UserIndex)
    
        Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
        'Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
        Dim MiObj As obj

        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

        End If
    
        'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
        ''If ObjData(MiObj.ObjIndex).Log = 1 Then
        '    Call LogDesarrollo(UserList(UserIndex).name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
        'End If
    
        Call SubirSkill(UserIndex, eSkill.Alquimia)
        Call UpdateUserInv(True, UserIndex, 0)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    End If

End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

    If Not UserList(UserIndex).Stats.MinSta > 0 Then
        Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If SastreTieneMateriales(UserIndex, ItemIndex) And _
        UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) >= ObjData(ItemIndex).SkMAGOria And _
        PuedeConstruirSastre(ItemIndex) And UserList(UserIndex).Invent.HerramientaEqpObjIndex = COSTURERO Then
        
        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        
        Call WriteUpdateSta(UserIndex)
    
        Call SastreQuitarMateriales(UserIndex, ItemIndex)
    
        ' If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        ' Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, " 1", 5))
        ' UserList(UserIndex).flags.UltimoMensaje = 9
        ' End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(63, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
        Dim MiObj As obj

        MiObj.Amount = 1
        MiObj.ObjIndex = ItemIndex

        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
    
        Call SubirSkill(UserIndex, eSkill.Herreria)
        Call UpdateUserInv(True, UserIndex, 0)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

    End If
    
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales, ByVal cant As Byte) As Integer

    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 50 * cant

        Case iMinerales.PlataCruda
            MineralesParaLingote = 70 * cant

        Case iMinerales.OroCrudo
            MineralesParaLingote = 90 * cant

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
    '    Call LogTarea("Sub DoLingotes")

    If UserList(UserIndex).Stats.MinSta > 5 Then
        Call QuitarSta(UserIndex, 5)
    
    Else
        
        Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub

    End If

    Dim slot As Integer
    Dim obji As Integer

    slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).Invent.Object(slot).ObjIndex
    
    Dim cant As Byte: cant = RandomNumber(1, 3)
    
    Dim necesarios As Integer

    necesarios = MineralesParaLingote(obji, cant)
    
    If UserList(UserIndex).Invent.Object(slot).Amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
        Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub

    End If
    
    UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount - MineralesParaLingote(obji, cant)

    If UserList(UserIndex).Invent.Object(slot).Amount < 1 Then
        UserList(UserIndex).Invent.Object(slot).Amount = 0
        UserList(UserIndex).Invent.Object(slot).ObjIndex = 0

    End If
    
    Dim nPos  As WorldPos

    Dim MiObj As obj

    MiObj.Amount = cant
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

    End If

    Call UpdateUserInv(False, UserIndex, slot)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateRenderValue(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, cant, 5))
    'If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
    '  Call WriteConsoleMsg(UserIndex, "¡Has obtenido lingotes!", FontTypeNames.FONTTYPE_INFO)
            
    '  UserList(UserIndex).flags.UltimoMensaje = 5
    'End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
    Call SubirSkill(UserIndex, eSkill.Herreria)
  
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
        
    If UserList(UserIndex).Counters.Trabajando = 1 And Not UserList(UserIndex).flags.UsandoMacro Then
        Call WriteMacroTrabajoToggle(UserIndex, True)

    End If
    
End Sub

Function ModFundicion(ByVal clase As eClass) As Single

    Select Case clase

        Case eClass.Trabajador
            ModFundicion = 3

        Case Else
            ModFundicion = 1

    End Select

End Function

Function ModAlquimia(ByVal clase As eClass) As Integer

    Select Case clase

        Case eClass.Druid
            ModAlquimia = 1

        Case eClass.Trabajador
            ModAlquimia = 1

        Case Else
            ModAlquimia = 3

    End Select

End Function

Function ModSastre(ByVal clase As eClass) As Integer

    Select Case clase

        Case eClass.Trabajador
            ModSastre = 1

        Case Else
            ModSastre = 3

    End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer

    Select Case clase

        Case eClass.Trabajador
            ModCarpinteria = 1

        Case Else
            ModCarpinteria = 3

    End Select

End Function

Function ModHerreriA(ByVal clase As eClass) As Single

    Select Case clase

        Case eClass.Trabajador
            ModHerreriA = 1

        Case Else
            ModHerreriA = 3

    End Select

End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
                
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.Body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Char.Body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))

End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

    Dim Suerte    As Byte

    Dim exito     As Byte

    Dim obj       As obj

    Dim posMadera As WorldPos

    If Not LegalPos(Map, x, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .x = x
        .Y = Y

    End With

    If MapData(Map, x, Y).ObjInfo.ObjIndex <> 58 Then
        Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
        Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
        ' Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "No podés hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, x, Y).ObjInfo.Amount < 3 Then
        Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
        Suerte = 3
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
        Suerte = 2
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
        Suerte = 1

    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        obj.ObjIndex = FOGATA_APAG
        obj.Amount = MapData(Map, x, Y).ObjInfo.Amount \ 3
    
        Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.Amount & " ramitas.", FontTypeNames.FONTTYPE_INFO)
    
        Call MakeObj(obj, Map, x, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    Else

        '[CDT 17-02-2004]
        If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
            Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 10

        End If

        '[/CDT]
    End If

    Call SubirSkill(UserIndex, Supervivencia)

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal RedDePesca As Boolean = False)

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
        
            MiObj.Amount = IIf(RedDePesca, RandomNumber(1, 3), 1)
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
                res = RandomNumber(1, IIf(RedDePesca, EspecialesPesca(i).Amount * 2, EspecialesPesca(i).Amount)) ' Red de pesca chance x2 (revisar)
            
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

    If Not MapInfo(UserList(victimaindex).Pos.Map).Seguro = 0 Then Exit Sub

    If UserList(LadrOnIndex).flags.Seguro Then
        Call WriteConsoleMsg(LadrOnIndex, "Debes quitar el seguro para robar", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub

    End If

    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then
        Call WriteConsoleMsg(LadrOnIndex, "Para robar debes salir de la armada real.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub

    End If

    If TriggerZonaPelea(LadrOnIndex, victimaindex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(LadrOnIndex).GuildIndex > 0 Then
        If UserList(LadrOnIndex).flags.SeguroClan Then
            If UserList(LadrOnIndex).GuildIndex = UserList(victimaindex).GuildIndex Then
                Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

        End If

    End If

    If UserList(LadrOnIndex).Grupo.EnGrupo > 0 Then
        If UserList(LadrOnIndex).GuildIndex = UserList(victimaindex).GuildIndex Then
            Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub

        End If

    End If

    If UserList(LadrOnIndex).Grupo.EnGrupo = True Then

        Dim i As Byte

        For i = 1 To UserList(UserList(LadrOnIndex).Grupo.Lider).Grupo.CantidadMiembros

            If UserList(UserList(LadrOnIndex).Grupo.Lider).Grupo.Miembros(i) = victimaindex Then
                Call WriteConsoleMsg(LadrOnIndex, "No podes robarle a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If

        Next i

    End If

    Call QuitarSta(LadrOnIndex, 15)

    If UserList(victimaindex).flags.Privilegios And PlayerType.user Then

        Dim Suerte     As Integer

        Dim res        As Integer

        Dim Porcentaje As Byte
    
        If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
            Suerte = 35
            Porcentaje = 1
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
            Suerte = 30
            Porcentaje = 1
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
            Suerte = 28
            Porcentaje = 2
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
            Suerte = 24
            Porcentaje = 3
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
            Suerte = 22
            Porcentaje = 4
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
            Suerte = 20
            Porcentaje = 5
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
            Suerte = 18
            Porcentaje = 6
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
            Suerte = 15
            Porcentaje = 7
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
            Suerte = 10
            Porcentaje = 8
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
            Suerte = 7
            Porcentaje = 9
        ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
            Suerte = 5
            Porcentaje = 10

        End If

        res = RandomNumber(1, Suerte)
        
        If res < 4 Then 'Exito robo
    
            ' TODO: Clase ladrón
            'If UserList(LadrOnIndex).clase = eClass.Trabajador Then
            If False Then
           
                If (RandomNumber(1, 50) < 25) Then
                    If TieneObjetosRobables(victimaindex) Then
                        Call RobarObjeto(LadrOnIndex, victimaindex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(victimaindex).Stats.GLD > 0 Then

                        Dim n As Long
                    
                        'porcentaje
                        n = UserList(victimaindex).Stats.GLD / 100 * Porcentaje
                    
                        ' N = RandomNumber(100, 1000)
                        If n > UserList(victimaindex).Stats.GLD Then n = UserList(victimaindex).Stats.GLD
                        UserList(victimaindex).Stats.GLD = UserList(victimaindex).Stats.GLD - n
                        
                        UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n

                        If UserList(LadrOnIndex).Stats.GLD > MAXORO Then UserList(LadrOnIndex).Stats.GLD = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(victimaindex).name, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Else

                If UserList(victimaindex).Stats.GLD > 0 Then

                    n = UserList(victimaindex).Stats.GLD / 100 * 0.5
                    
                    If n > UserList(victimaindex).Stats.GLD Then n = UserList(victimaindex).Stats.GLD
                    UserList(victimaindex).Stats.GLD = UserList(victimaindex).Stats.GLD - n
                        
                    UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + n

                    If UserList(LadrOnIndex).Stats.GLD > MAXORO Then UserList(LadrOnIndex).Stats.GLD = MAXORO
                        
                    Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(victimaindex).name, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If
        
        End If
    
    Else
        Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(victimaindex, "¡" & UserList(LadrOnIndex).name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(victimaindex, "¡" & UserList(LadrOnIndex).name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(victimaindex)

    End If

    If Status(LadrOnIndex) = 1 Then
        Call VolverCriminal(LadrOnIndex)

    End If
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    Call SubirSkill(LadrOnIndex, Robar)
    Call WriteUpdateGold(LadrOnIndex)
    Call WriteUpdateGold(victimaindex)

End Sub

Public Function ObjEsRobable(ByVal victimaindex As Integer, ByVal slot As Integer) As Boolean
    ' Agregué los barcos
    ' Esta funcion determina qué objetos son robables.

    Dim OI As Integer

    OI = UserList(victimaindex).Invent.Object(slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(victimaindex).Invent.Object(slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).donador = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otRunas And ObjData(OI).Instransferible = 0 And ObjData(OI).OBJType <> eOBJType.otMonturas

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)

    'Call LogTarea("Sub RobarObjeto")
    Dim flag As Boolean

    Dim i    As Integer

    flag = False

    If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
        i = 1

        Do While Not flag And i <= MAX_INVENTORY_SLOTS

            'Hay objeto en este slot?
            If UserList(victimaindex).Invent.Object(i).ObjIndex > 0 Then
                If ObjEsRobable(victimaindex, i) Then
                    If RandomNumber(1, 10) < 4 Then flag = True

                End If

            End If

            If Not flag Then i = i + 1
        Loop
    Else
        i = 20

        Do While Not flag And i > 0

            'Hay objeto en este slot?
            If UserList(victimaindex).Invent.Object(i).ObjIndex > 0 Then
                If ObjEsRobable(victimaindex, i) Then
                    If RandomNumber(1, 10) < 4 Then flag = True

                End If

            End If

            If Not flag Then i = i - 1
        Loop

    End If

    If flag Then

        Dim MiObj As obj

        Dim num   As Byte

        'Cantidad al azar
        num = RandomNumber(1, 5)
                
        If num > UserList(victimaindex).Invent.Object(i).Amount Then
            num = UserList(victimaindex).Invent.Object(i).Amount

        End If
                
        MiObj.Amount = num
        MiObj.ObjIndex = UserList(victimaindex).Invent.Object(i).ObjIndex
    
        UserList(victimaindex).Invent.Object(i).Amount = UserList(victimaindex).Invent.Object(i).Amount - num
                
        If UserList(victimaindex).Invent.Object(i).Amount <= 0 Then
            Call QuitarUserInvItem(victimaindex, CByte(i), 1)

        End If
            
        Call UpdateUserInv(False, victimaindex, CByte(i))
                
        If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

        End If
    
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)

    Else
        Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)

    End If

    'If exiting, cancel de quien es robado
    Call CancelExit(victimaindex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el daño
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer
    
    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)
    
    Select Case UserList(UserIndex).clase

        Case eClass.Assasin '35
            Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 4.25)
        
        Case eClass.Cleric, eClass.Paladin ' 15
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
        
        Case eClass.Bard, eClass.Druid '13
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
        
        Case Else '8
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select
    
    If RandomNumber(0, 70) < Suerte Then
        If VictimUserIndex <> 0 Then
            If UserList(UserIndex).clase = eClass.Assasin Then
                daño = Round(daño * 1.4, 0)
            Else
                daño = Round(daño * 1.2, 0)

            End If
            
            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño

            Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageEfectOverHead("¡" & daño & "!", UserList(VictimUserIndex).Char.CharIndex, &HFFFF00))

            If UserList(UserIndex).ChatCombate = 1 Then
                'Call WriteEfectOverHead(UserIndex, daño, UserList(UserIndex).Char.CharIndex) 'LADDER 21.11.08
                Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

            End If

            If UserList(VictimUserIndex).ChatCombate = 1 Then
                Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call FlushBuffer(VictimUserIndex)
        Else
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Int(daño * 1.5)

            If UserList(UserIndex).ChatCombate = 1 Then
                'Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 1.5), FontTypeNames.FONTTYPE_FIGHT)
                Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, Int(daño * 1.5))

            End If
            
            Call SendData(SendTarget.ToNPCArea, VictimNpcIndex, PrepareMessageEfectOverHead("¡" & daño * 1.5 & "!", Npclist(VictimNpcIndex).Char.CharIndex, &HFFFF00))

            '[Alejo]
            Call CalcularDarExp(UserIndex, VictimNpcIndex, Int(daño * 1.5))

        End If

    Else

        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)

        End If

    End If
    
    Call SubirSkill(UserIndex, Apuñalar)

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 28/01/2007
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer

    'If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
    If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).name <> "Espada Vikinga" Then Exit Sub

    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

    If RandomNumber(0, 100) < Suerte Then
        daño = Int(daño * 0.5)

        If VictimUserIndex <> 0 Then
            UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
            Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Else
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            '[Alejo]
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)

        End If

    End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub
    Call WriteUpdateSta(UserIndex)

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

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte)

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
    
            MiObj.Amount = RandomNumber(4, 7)
            
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
                res = RandomNumber(1, EspecialesTala(i).Amount)
                
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

Public Sub DoMineria(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte)

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
            
            MiObj.Amount = RandomNumber(2, 3)
    
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

    Dim Suerte       As Integer

    Dim res          As Integer

    Dim cant         As Integer

    Dim TActual      As Long

    Dim MeditarSkill As Byte

    With UserList(UserIndex)
        '.Counters.IdleCount = 0

        If .Stats.MinMAN >= .Stats.MaxMAN Then Exit Sub
    
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 Then
            Suerte = 7
        Else
            Suerte = 5

        End If
    
        If .flags.RegeneracionMana = 1 Then
            Suerte = 10
        End If
        
        res = RandomNumber(1, Suerte)
    
        If res = 1 Then
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)

            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
            '  If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
            '     Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
            '     UserList(UserIndex).flags.UltimoMensaje = 22
            '  End If
            
            Call WriteUpdateMana(UserIndex)
            Call SubirSkill(UserIndex, Meditar)

        End If

    End With

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

    Dim Suerte As Integer

    Dim res    As Integer

    If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
        Suerte = 35
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
        Suerte = 30
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
        Suerte = 28
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
        Suerte = 24
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
        Suerte = 22
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
        Suerte = 20
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
        Suerte = 18
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
        Suerte = 15
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
        Suerte = 10
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
        Suerte = 7
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
        Suerte = 5

    End If

    res = RandomNumber(1, Suerte)

    If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

        If UserList(VictimIndex).Stats.ELV < 20 Then
            Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

        End If

        Call FlushBuffer(VictimIndex)

    End If

End Sub

Public Sub DoMontar(ByVal UserIndex As Integer, ByRef Montura As ObjData, ByVal slot As Integer)

    If Not CheckRazaTipo(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then
        Call WriteConsoleMsg(UserIndex, "Tu raza no te permite usar esta montura.", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If Not CheckClaseTipo(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then
        Call WriteConsoleMsg(UserIndex, "Tu clase no te permite usar esta montura.", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(UserIndex).Stats.UserSkills(eSkill.equitacion) < Montura.MinSkill Then
        'Call WriteConsoleMsg(UserIndex, "No tenes suficientes conocimientos para usar esta montura.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitacion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    'Ladder 21/11/08
    If UserList(UserIndex).flags.Montado = 0 Then
        If (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger > 10) Then
            Call WriteConsoleMsg(UserIndex, "No podés montar aquí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

    End If

    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        Call WriteLocaleMsg(UserIndex, "123", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).Char.FX = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
    End If

    If UserList(UserIndex).flags.Montado = 1 Then
        If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
                UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica
                Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MonturaSlot)

            End If

        End If

    End If

    UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
    UserList(UserIndex).Invent.MonturaSlot = slot

    If UserList(UserIndex).flags.Montado = 0 Then

        If ObjData(UserList(UserIndex).Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
            UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + Montura.ResistenciaMagica

        End If
            
        If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
            'UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).RopajeBajo
            UserList(UserIndex).Char.Body = Montura.RopajeBajo
        Else
            UserList(UserIndex).Char.Body = Montura.Ropaje

        End If

        'UserList(UserIndex).Char.body = Montura.Ropaje
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).Char.CascoAnim
        UserList(UserIndex).flags.Montado = 1
        UserList(UserIndex).Char.speeding = VelocidadMontura
    Else
        UserList(UserIndex).flags.Montado = 0
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
        UserList(UserIndex).Char.speeding = VelocidadNormal

        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        
            If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).RopajeBajo
            Else
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

            End If

        Else
            Call DarCuerpoDesnudo(UserIndex)

        End If
            
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim

        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

    End If

    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

    Call UpdateUserInv(False, UserIndex, slot)
    Call WriteEquiteToggle(UserIndex)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

End Sub

Public Function ApuñalarFunction(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer) As Integer

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el daño
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer

    Dim Random As Byte

    Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

    Select Case UserList(UserIndex).clase

        Case eClass.Assasin '35
            Suerte = Int(((0.00003 * Skill - 0.001) * Skill + 0.098) * Skill + 4.25)
        
            If VictimNpcIndex = 0 Then
                If UserList(VictimUserIndex).Char.heading = UserList(UserIndex).Char.heading Then
                    Random = RandomNumber(1, 3)

                    If Random = 1 Then
                        Suerte = 70

                    End If

                End If

            End If
    
        Case eClass.Cleric, eClass.Paladin ' 15
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bard, eClass.Druid '13
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else '8
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    If UserList(UserIndex).clase = eClass.Assasin Then
        daño = Round(daño * 0.4, 0)
    Else
        daño = Round(daño * 0.1, 0)

    End If

    If RandomNumber(0, 70) < Suerte Then
        If VictimUserIndex <> 0 Then

            ApuñalarFunction = daño
        
            If UserList(UserIndex).ChatCombate = 1 Then
                Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            If UserList(VictimUserIndex).ChatCombate = 1 Then
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).name & " te ha apuñalado por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            'Call SendData(SendTarget.ToAll, 0, PrepareMessageEfectToScreen(&HFF, 350))           'Rayo
        
            Call WriteEfectToScreen(VictimUserIndex, &H3C3CFF, 200, True)
            Call WriteEfectToScreen(UserIndex, &H3C3CFF, 150, True)
        
            Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateFX(UserList(VictimUserIndex).Char.CharIndex, 89, 0))
        
            Call FlushBuffer(VictimUserIndex)
            Call FlushBuffer(UserIndex)
        
        Else
            Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
            '[Alejo]
        
            ApuñalarFunction = daño
        
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)

        End If

    End If

End Function

Public Sub ActualizarRecurso(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)

    Dim ObjIndex As Integer

    ObjIndex = MapData(Map, x, Y).ObjInfo.ObjIndex

    Dim TiempoActual As Long

    TiempoActual = timeGetTime

    ' Data = Ultimo uso
    If (TiempoActual - MapData(Map, x, Y).ObjInfo.data) * 0.001 > ObjData(ObjIndex).TiempoRegenerar Then
        MapData(Map, x, Y).ObjInfo.Amount = ObjData(ObjIndex).VidaUtil
        MapData(Map, x, Y).ObjInfo.data = &H7FFFFFFF   ' Ultimo uso = Max Long

    End If

End Sub

Public Function ObtenerPezRandom(ByVal PoderCania As Integer) As Long

    Dim i As Long, SumaPesos As Long, ValorGenerado As Long
    
    If PoderCania > UBound(PesoPeces) Then PoderCania = UBound(PesoPeces)
    SumaPesos = PesoPeces(PoderCania)
    
    ValorGenerado = RandomNumber(1, SumaPesos)
    
    ObtenerPezRandom = Peces(BinarySearchPeces(ValorGenerado)).ObjIndex
    
End Function
