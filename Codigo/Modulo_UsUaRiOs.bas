Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)

    Dim DaExp       As Integer

    Dim EraCriminal As Byte
    
    DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)
    
    If UserList(attackerIndex).Stats.ELV < STAT_MAXELV Then
        UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp

        If UserList(attackerIndex).Stats.Exp > MAXEXP Then UserList(attackerIndex).Stats.Exp = MAXEXP

        Call WriteUpdateExp(attackerIndex)
        Call CheckUserLevel(attackerIndex)

    End If
    
    'Lo mata
    'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
    
    Call WriteLocaleMsg(attackerIndex, "184", FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).name)
    Call WriteLocaleMsg(attackerIndex, "140", FontTypeNames.FONTTYPE_EXP, DaExp)
          
    'Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteLocaleMsg(VictimIndex, "185", FontTypeNames.FONTTYPE_FIGHT, UserList(attackerIndex).name)
    
    If TriggerZonaPelea(VictimIndex, attackerIndex) <> TRIGGER6_PERMITE Then
        EraCriminal = Status(attackerIndex)
        
        If EraCriminal = 2 And Status(attackerIndex) < 2 Then
            Call RefreshCharStatus(attackerIndex)
        ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
            Call RefreshCharStatus(attackerIndex)

        End If

    End If
    
    Call UserDie(VictimIndex)
    
    If UserList(attackerIndex).flags.BattleModo = 1 Then
        Call ContarPuntoBattle(VictimIndex, attackerIndex)

    End If
    
    If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1
    'Call CheckearRecompesas(attackerIndex, 2)
    
    Call FlushBuffer(VictimIndex)

End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)

    UserList(UserIndex).Char.speeding = VelocidadNormal
    'Call WriteVelocidadToggle(UserIndex)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)

    If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp

    End If

    If UserList(UserIndex).flags.Navegando = 1 Then

        Dim Barco As ObjData

        Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)

        If Barco.Ropaje <> iTraje Then
            UserList(UserIndex).Char.Head = 0
            UserList(UserIndex).Char.CascoAnim = NingunCasco

        End If
    
        If Barco.Ropaje = iTraje Then
            UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
            Call WriteNadarToggle(UserIndex, True)
        
        Else
            Call WriteNadarToggle(UserIndex, False)
        
        End If
    
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
    
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        ' UserList(UserIndex).Char.CascoAnim = NingunCasco
   
        UserList(UserIndex).Char.speeding = Barco.Velocidad
        'Call WriteVelocidadToggle(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))
   
    Else

        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).RopajeBajo
            Else
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

            End If

        Else
            Call DarCuerpoDesnudo(UserIndex)

        End If
    
        UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head

        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

        End If

        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

        End If

        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
    
            If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CreaGRH <> "" Then
                UserList(UserIndex).Char.Arma_Aura = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 1))

            End If
        
        End If

        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).RopajeBajo
            Else
                UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje

            End If
    
            If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).CreaGRH <> "" Then
                UserList(UserIndex).Char.Arma_Aura = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 2))

            End If
        
        End If

        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

            If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).CreaGRH <> "" Then
                UserList(UserIndex).Char.Escudo_Aura = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Escudo_Aura, False, 3))

            End If
        
        End If

        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

            If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CreaGRH <> "" Then
                UserList(UserIndex).Char.Head_Aura = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Head_Aura, False, 4))

            End If
        
        End If

        If UserList(UserIndex).Invent.MagicoObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CreaGRH <> "" Then
                UserList(UserIndex).Char.Otra_Aura = ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Otra_Aura, False, 5))

            End If

        End If

        If UserList(UserIndex).Invent.NudilloObjIndex > 0 Then
            If ObjData(UserList(UserIndex).Invent.NudilloObjIndex).CreaGRH <> "" Then
                UserList(UserIndex).Char.Arma_Aura = ObjData(UserList(UserIndex).Invent.NudilloObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 1))

            End If

        End If

    End If

    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

    Call WriteUpdateUserStats(UserIndex)

    Call WriteHora(UserIndex)

End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As Byte, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    With UserList(UserIndex).Char
        .Body = Body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = Casco

    End With
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, Head, heading, UserList(UserIndex).Char.CharIndex, Arma, Escudo, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, Casco))

End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, ByVal Desvanecer As Boolean)

    On Error GoTo ErrorHandler

    Dim Error As String
   
    Error = "1"

    If UserList(UserIndex).Char.CharIndex = 0 Then Exit Sub
   
    CharList(UserList(UserIndex).Char.CharIndex) = 0
    
    If UserList(UserIndex).Char.CharIndex = LastChar Then

        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1

            If LastChar <= 1 Then Exit Do
        Loop

    End If

    Error = "2"
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex, Desvanecer))
    Error = "3"
    Call QuitarUser(UserIndex, UserList(UserIndex).Pos.Map)
    Error = "4"
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).UserIndex = 0
    Error = "5"
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    Error = "6"
    Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Error & " - " & Err.Number & ": " & Err.description)

End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)

    '*************************************************
    'Author: Tararira
    'Last modified: 6/04/2007
    'Refreshes the status and tag of UserIndex.
    '*************************************************
    Dim klan As String

    If UserList(UserIndex).GuildIndex > 0 Then
        klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
        klan = " <" & klan & ">"

    End If
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.Status, UserList(UserIndex).name & klan))

End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal appear As Byte = 0)

    On Error GoTo hayerror

    Dim CharIndex As Integer

    Dim errort    As String

    If InMapBounds(Map, x, Y) Then

        'If needed make a new character in list
        If UserList(UserIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
            
        End If

        errort = "1"
        
        'Place character on map if needed
        If toMap Then MapData(Map, x, Y).UserIndex = UserIndex
        errort = "2"
        
        'Send make character command to clients
        Dim klan       As String

        Dim clan_nivel As Byte

        If UserList(UserIndex).GuildIndex > 0 Then
            klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
            clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

        End If

        errort = "3"
        
        Dim bCr As Byte
        
        bCr = UserList(UserIndex).Faccion.Status
        errort = "4"
        
        If LenB(klan) <> 0 Then
            If Not toMap Then
                errort = "5"
                Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, x, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name & " <" & klan & ">", bCr, UserList(UserIndex).flags.Privilegios, UserList(UserIndex).Char.ParticulaFx, UserList(UserIndex).Char.Head_Aura, UserList(UserIndex).Char.Arma_Aura, UserList(UserIndex).Char.Body_Aura, UserList(UserIndex).Char.Otra_Aura, UserList(UserIndex).Char.Escudo_Aura, UserList(UserIndex).Char.speeding, False, UserList(UserIndex).donador.activo, appear, UserList(UserIndex).Grupo.Lider, UserList(UserIndex).GuildIndex, clan_nivel, UserList(UserIndex).Stats.MinHp, UserList(UserIndex).Stats.MaxHp, 0)
            Else
                errort = "6"
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map, appear)

            End If

        Else 'if tiene clan

            If Not toMap Then
                errort = "7"
                Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.CharIndex, x, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name, bCr, UserList(UserIndex).flags.Privilegios, UserList(UserIndex).Char.ParticulaFx, UserList(UserIndex).Char.Head_Aura, UserList(UserIndex).Char.Arma_Aura, UserList(UserIndex).Char.Body_Aura, UserList(UserIndex).Char.Otra_Aura, UserList(UserIndex).Char.Escudo_Aura, UserList(UserIndex).Char.speeding, False, UserList(UserIndex).donador.activo, appear, UserList(UserIndex).Grupo.Lider, 0, 0, UserList(UserIndex).Stats.MinHp, UserList(UserIndex).Stats.MaxHp, 0)
            Else
                errort = "8"
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map, appear)

            End If

        End If 'if clan

    End If

    Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description & " - Nombre del usuario " & UserList(UserIndex).name) & " - " & errort & "- Pos: M: " & Map & " X: " & x & " Y: " & Y
    'Resume Next
    Call CloseSocket(UserIndex)

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 01/10/2007
    'Chequea que el usuario no halla alcanzado el siguiente nivel,
    'de lo contrario le da la vida, mana, etc, correspodiente.
    '07/08/2006 Integer - Modificacion de los valores
    '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
    '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
    '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
    '*************************************************

    On Error GoTo Errhandler

    Dim Pts              As Integer

    Dim AumentoHIT       As Integer

    Dim AumentoMANA      As Integer

    Dim AumentoSTA       As Integer

    Dim AumentoHP        As Integer

    Dim WasNewbie        As Boolean

    Dim Promedio         As Double

    Dim aux              As Integer

    Dim DistVida(1 To 5) As Integer
    
    Dim PasoDeNivel      As Boolean
    
    With UserList(UserIndex)
    
        '¿Alcanzo el maximo nivel?
        If .Stats.ELV >= STAT_MAXELV Then
            .Stats.Exp = 0
            .Stats.ELU = 0
            Exit Sub

        End If
            
        WasNewbie = EsNewbie(UserIndex)
        
        Do While .Stats.Exp >= .Stats.ELU
            
            'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub

            End If
            
            'Store it!
            'Call Statistics.UserLevelUp(UserIndex)

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 106, 0))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.x, .Pos.Y))
            Call WriteLocaleMsg(UserIndex, "186", FontTypeNames.FONTTYPE_INFO)
            
            Pts = Pts + 5
            
            .Stats.ELV = .Stats.ELV + 1
            
            .Stats.Exp = .Stats.Exp - .Stats.ELU

            If .Stats.ELV < 15 Then
                .Stats.ELU = .Stats.ELU * 1.4
            ElseIf .Stats.ELV < 21 Then
                .Stats.ELU = .Stats.ELU * 1.35
            ElseIf .Stats.ELV < 26 Then
                .Stats.ELU = .Stats.ELU * 1.3
            ElseIf .Stats.ELV < 35 Then
                .Stats.ELU = .Stats.ELU * 1.2
            ElseIf .Stats.ELV < 40 Then
                .Stats.ELU = .Stats.ELU * 1.3
            Else
                .Stats.ELU = .Stats.ELU * 1.375

            End If
            
            'Calculo subida de vida
            Promedio = ModVida(.clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)
            
            If Promedio - Int(Promedio) = 0.5 Then
                'Es promedio semientero
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 1.5

                End If

            Else
                'Es promedio entero
                DistVida(1) = DistribucionEnteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 2
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(4) Then
                    AumentoHP = Promedio - 1
                Else
                    AumentoHP = Promedio - 2

                End If
                
            End If
            
            Select Case .clase

                Case eClass.Mage '
                    AumentoHIT = 1
                    AumentoMANA = 2.8 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago

                Case eClass.Bard 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef

                Case eClass.Druid 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef

                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef

                Case eClass.Cleric 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef

                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef

                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef

                Case eClass.Trabajador
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef + 5

                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef

                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef

            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHp = .Stats.MaxHp + AumentoHP

            If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA

            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

            'Actualizamos Golpe Máximo
            .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            
            'Actualizamos Golpe Mínimo
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
        
            'Notificamos al user
            If AumentoHP > 0 Then
                'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "197", FontTypeNames.FONTTYPE_INFO, AumentoHP)

            End If

            If AumentoSTA > 0 Then
                'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "198", FontTypeNames.FONTTYPE_INFO, AumentoSTA)

            End If

            If AumentoMANA > 0 Then
                Call WriteLocaleMsg(UserIndex, "199", FontTypeNames.FONTTYPE_INFO, AumentoMANA)

                'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
            End If

            If AumentoHIT > 0 Then
                Call WriteLocaleMsg(UserIndex, "200", FontTypeNames.FONTTYPE_INFO, AumentoHIT)

                'Call WriteConsoleMsg(UserIndex, "Tu golpe aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
            End If

            PasoDeNivel = True
             
            ' Call LogDesarrollo(.name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)

            .Stats.MinHp = .Stats.MaxHp
            
            ' Call UpdateUserInv(True, UserIndex, 0)
            
            If EsNewbie(UserIndex) Then

                Dim OroRecompenza As Long

                OroRecompenza = OroPorNivel * .Stats.ELV * OroMult * .flags.ScrollOro
                .Stats.GLD = .Stats.GLD + OroRecompenza
                'Call WriteConsoleMsg(UserIndex, "Has ganado " & OroRecompenza & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "29", FontTypeNames.FONTTYPE_INFO, OroRecompenza)
                
            End If
            
            If Not EsNewbie(UserIndex) And WasNewbie Then
        
                Call QuitarNewbieObj(UserIndex)
            
            End If
        
        Loop
        
        If PasoDeNivel Then
            Call UpdateUserInv(True, UserIndex, 0)
            'Call CheckearRecompesas(UserIndex, 3)
            Call WriteUpdateUserStats(UserIndex)
            
            If Pts > 0 Then
                
                .Stats.SkillPts = .Stats.SkillPts + Pts
                Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                Call WriteLocaleMsg(UserIndex, "187", FontTypeNames.FONTTYPE_INFO, Pts)

                'Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
            End If

        End If
    
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean

    PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1
    'If PuedeAtravesarAgua = True Then
    '  Exit Function
    'Else
    '  If UserList(UserIndex).flags.Nadando = 1 Then
    ' PuedeAtravesarAgua = True
    'End If
    'End If

End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)

    Dim nPos         As WorldPos

    Dim nPosOriginal As WorldPos

    Dim nPosMuerto   As WorldPos

    Dim sailing      As Boolean

    With UserList(UserIndex)

        If .accion.AccionPendiente = True Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, .accion.Particula, 1, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 1, Accion_Barra.CancelarAccion))
            .accion.AccionPendiente = False
            .accion.Particula = 0
            .accion.TipoAccion = Accion_Barra.CancelarAccion
            .accion.HechizoPendiente = 0
            .accion.RunaObj = 0
            .accion.ObjSlot = 0
            .accion.AccionPendiente = False

        End If

        sailing = PuedeAtravesarAgua(UserIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)

    End With
        
    If MapData(nPos.Map, nPos.x, nPos.Y).TileExit.Map <> 0 And UserList(UserIndex).Counters.TiempoDeMapeo > 0 Then
        If UserList(UserIndex).flags.Muerto = 0 Then
            Call WriteConsoleMsg(UserIndex, "Estas en combate, debes aguardar " & UserList(UserIndex).Counters.TiempoDeMapeo & " segundo(s) para escapar...", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WritePosUpdate(UserIndex)
            Exit Sub

        End If

    End If

    If MapData(nPos.Map, nPos.x, nPos.Y).UserIndex <> 0 Then
        If UserList(MapData(nPos.Map, nPos.x, nPos.Y).UserIndex).flags.Muerto = 1 Then

            Dim IndexMuerto As Integer

            IndexMuerto = MapData(nPos.Map, nPos.x, nPos.Y).UserIndex

            With UserList(IndexMuerto)
                'Call WarpToLegalPos(IndexMuerto, .Pos.Map, .Pos.X, .Pos.Y, False)
                    
                If .accion.AccionPendiente = True Then
                    Call SendData(SendTarget.ToPCArea, IndexMuerto, PrepareMessageParticleFX(.Char.CharIndex, .accion.Particula, 1, True))
                    Call SendData(SendTarget.ToPCArea, IndexMuerto, PrepareMessageBarFx(.Char.CharIndex, 1, Accion_Barra.CancelarAccion))
                    .accion.AccionPendiente = False
                    .accion.Particula = 0
                    .accion.TipoAccion = Accion_Barra.CancelarAccion
                    .accion.HechizoPendiente = 0
                    .accion.RunaObj = 0
                    .accion.ObjSlot = 0
                    .accion.AccionPendiente = False

                    'Call WritePosUpdate(UserIndex)
                    'Call WarpToLegalPos(IndexMuerto, .Pos.Map, .Pos.X, .Pos.Y)
                End If

                Call WarpToLegalPos(IndexMuerto, .Pos.Map, .Pos.x, .Pos.Y, False)
                    
            End With

        Else
            Call WritePosUpdate(UserIndex)

            'Call WritePosUpdate(MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
        End If

    End If

    If LegalPos(UserList(UserIndex).Pos.Map, nPos.x, nPos.Y, sailing, Not sailing, UserList(UserIndex).flags.Montado) Then
        If MapInfo(UserList(UserIndex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...

            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, nPos.x, nPos.Y))
        
        End If

        'Call RefreshAllUser(UserIndex) '¿Clones? Ladder probar
        'Update map and user pos
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).UserIndex = 0
        UserList(UserIndex).Pos = nPos
        UserList(UserIndex).Char.heading = nHeading
        MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).UserIndex = UserIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)
       
    Else
        Call WritePosUpdate(UserIndex)

    End If
    
    If UserList(UserIndex).Counters.Trabajando Then
        Call WriteMacroTrabajoToggle(UserIndex, False)

    End If

    If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

End Sub

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal slot As Byte, ByRef Object As UserOBJ)
    UserList(UserIndex).Invent.Object(slot) = Object
    Call WriteChangeInventorySlot(UserIndex, slot)

End Sub

Function NextOpenCharIndex() As Integer

    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS

        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then LastChar = LoopC
            
            Exit Function

        End If

    Next LoopC

End Function

Function NextOpenUser() As Integer

    Dim LoopC As Long
   
    For LoopC = 1 To MaxUsers + 1

        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
   
    NextOpenUser = LoopC

End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    Dim GuildI As Integer

    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHp & "/" & UserList(UserIndex).Stats.MaxHp & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit, FontTypeNames.FONTTYPE_INFO)

    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

    End If
    
    If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

    End If
    
    GuildI = UserList(UserIndex).GuildIndex

    If GuildI > 0 Then
        Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)

        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
            Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)

        End If

        'guildpts no tienen objeto
    End If
    
    #If ConUpTime Then

        Dim TempDate As Date

        Dim TempSecs As Long

        Dim TempStr  As String

        TempDate = Now - UserList(UserIndex).LogOnTime
        TempSecs = (UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
    #End If
    
    Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.x & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Veces que Moriste: " & UserList(UserIndex).flags.VecesQueMoriste, FontTypeNames.FONTTYPE_INFO)

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Shows the users Stats when the user is online.
    '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
    '*************************************************
    With UserList(UserIndex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos Matados: " & .Faccion.CiudadanosMatados & " Criminales Matados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)

        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)

        End If

        Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & .Stats.GLD, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en banco: " & .Stats.Banco, FontTypeNames.FONTTYPE_INFO)
    
        Call WriteConsoleMsg(sendIndex, "Cuenta: " & .Cuenta, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Creditos: " & .donador.CreditoDonador, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Fecha Vencimiento Donador: " & .donador.FechaExpiracion, FontTypeNames.FONTTYPE_INFO)
    
    End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Shows the users Stats when the user is offline.
    '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
    '*************************************************
    Dim CharFile      As String

    Dim Ban           As String

    Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & GetVar(CharFile, "STATS", "GLD"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en boveda: " & GetVar(CharFile, "STATS", "BANCO"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Cuenta: " & GetVar(CharFile, "INIT", "Cuenta"), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)

        End If
        
        Ban = GetVar(CharFile, "BAN", "BanMotivo")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)

        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason"), FontTypeNames.FONTTYPE_INFO)

        End If
        
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & CharName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim j As Long
    
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To UserList(UserIndex).CurrentInventorySlots

        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

        End If

    Next j

End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    On Error Resume Next

    Dim j        As Long

    Dim CharFile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)

    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim j As Integer

    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)

    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next
    Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

End Sub

Function DameUserIndex(SocketId As Integer) As Integer

    Dim LoopC As Integer
  
    LoopC = 1
  
    Do Until UserList(LoopC).ConnID = SocketId

        LoopC = LoopC + 1
    
        If LoopC > MaxUsers Then
            DameUserIndex = 0
            Exit Function

        End If
    
    Loop
  
    DameUserIndex = LoopC

End Function

Function DameUserIndexConNombre(ByVal nombre As String) As Integer

    Dim LoopC As Integer
  
    LoopC = 1
  
    nombre = UCase$(nombre)

    Do Until UCase$(UserList(LoopC).name) = nombre

        LoopC = LoopC + 1
    
        If LoopC > MaxUsers Then
            DameUserIndexConNombre = 0
            Exit Function

        End If
    
    Loop
  
    DameUserIndexConNombre = LoopC

End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    '**********************************************
    'Author: Unknown
    'Last Modification: 24/07/2007
    '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
    '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
    '**********************************************
    Dim EraCriminal As Byte

    'Guardamos el usuario que ataco el npc.

    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        Npclist(NpcIndex).Target = UserIndex
        Npclist(NpcIndex).Movement = TipoAI.NpcMaloAtacaUsersBuenos
        Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).name
    Else
    
    End If

    'Npc que estabas atacando.
    Dim LastNpcHit As Integer

    LastNpcHit = UserList(UserIndex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(UserIndex).flags.NPCAtacado = NpcIndex

    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then

        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).name Then

        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

    End If

    '  EraCriminal = Status(UserIndex)

    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If Status(UserIndex) = 1 Or Status(UserIndex) = 3 Then
            Call VolverCriminal(UserIndex)

        End If

    End If

End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        PuedeApuñalar = ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UserList(UserIndex).clase = eClass.Assasin) And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
    Else
        PuedeApuñalar = False

    End If

End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)

    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

    If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
        
        Dim Lvl As Integer

        Lvl = UserList(UserIndex).Stats.ELV
        
        If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
            
        If UserList(UserIndex).Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
        
        Dim Aumenta As Integer

        Dim Prob    As Integer

        Dim Menor   As Byte

        Menor = 10
             
        If Lvl <= 3 Then
            Prob = 25
        ElseIf Lvl > 3 And Lvl < 6 Then
            Prob = 35
        ElseIf Lvl >= 6 And Lvl < 10 Then
            Prob = 40
        ElseIf Lvl >= 10 And Lvl < 20 Then
            Prob = 45
        Else
            Prob = 50

        End If
             
        Aumenta = RandomNumber(1, Prob)
             
        If UserList(UserIndex).flags.PendienteDelExperto = 1 Then
            Menor = 15

        End If
        
        If Aumenta < Menor Then
            UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
    
            Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
            
            Dim BonusExp As Long

            BonusExp = 5 * ExpMult * UserList(UserIndex).flags.ScrollExp
        
            If UserList(UserIndex).donador.activo = 1 Then
                BonusExp = BonusExp * 1.1

            End If
        
            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + BonusExp

                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
            
                If UserList(UserIndex).ChatCombate = 1 Then
                    Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, BonusExp)

                End If
                
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)

            End If

        End If

    End If

End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal UserIndex As Integer)

    '************************************************
    'Author: Uknown
    'Last Modified: 04/15/2008 (NicoNZ)
    'Ahora se resetea el counter del invi
    '************************************************
    On Error GoTo ErrorHandler

    Dim i  As Long

    Dim aN As Integer
    
    'Sonido
        
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MUERTE_HOMBRE, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    UserList(UserIndex).Stats.MinHp = 0
    UserList(UserIndex).Stats.MinSta = 0
    UserList(UserIndex).flags.AtacadoPorUser = 0
    UserList(UserIndex).flags.Envenenado = 0
    UserList(UserIndex).flags.Ahogandose = 0
    UserList(UserIndex).flags.Incinerado = 0
    UserList(UserIndex).flags.incinera = 0
    UserList(UserIndex).flags.Paraliza = 0
    UserList(UserIndex).flags.Envenena = 0
    UserList(UserIndex).flags.Estupidiza = 0
    UserList(UserIndex).flags.Muerto = 1
    'UserList(UserIndex).flags.SeguroParty = True
    'Call WritePartySafeOn(UserIndex)
    
    aN = UserList(UserIndex).flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString

    End If
    
    aN = UserList(UserIndex).flags.NPCAtacado

    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString

        End If

    End If

    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.NPCAtacado = 0
    
    '<<<< Paralisis >>>>
    If UserList(UserIndex).flags.Paralizado = 1 Then
        UserList(UserIndex).flags.Paralizado = 0
        Call WriteParalizeOK(UserIndex)

    End If
    
    '<<<< Inmovilizado >>>>
    If UserList(UserIndex).flags.Inmovilizado = 1 Then
        UserList(UserIndex).flags.Inmovilizado = 0
        Call WriteInmovilizaOK(UserIndex)

    End If
    
    '<<< Estupidez >>>
    If UserList(UserIndex).flags.Estupidez = 1 Then
        UserList(UserIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)

    End If
    
    '<<<< Descansando >>>>
    If UserList(UserIndex).flags.Descansar Then
        UserList(UserIndex).flags.Descansar = False
        Call WriteRestOK(UserIndex)

    End If
    
    '<<<< Meditando >>>>
    If UserList(UserIndex).flags.Meditando Then
        UserList(UserIndex).flags.Meditando = False
        UserList(UserIndex).Char.FX = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
    End If
    
    'If UserList(UserIndex).Familiar.Invocado = 1 Then
    ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso("17", Npclist(UserList(UserIndex).Familiar.Id).Pos.x, Npclist(UserList(UserIndex).Familiar.Id).Pos.Y))
    ' UserList(UserIndex).Familiar.Invocado = 0
    ' Call QuitarNPC(UserList(UserIndex).Familiar.Id)
    ' End If
    
    '<<<< Invisible >>>>
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Counters.Invisibilidad = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))

    End If

    If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 Then '  Ladder 06/07/2014 Si el mapa es seguro, no se caen los items
        If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then

            ' << Si es newbie no pierde el inventario >>
            If UserList(UserIndex).flags.Privilegios = user Then
        
                If Not EsNewbie(UserIndex) Then
                    If UserList(UserIndex).flags.PendienteDelSacrificio = 0 Then
                
                        Call TirarTodo(UserIndex)
                    
                        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.ArmourEqpObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)

                            End If

                        End If

                        'desequipar arma
                        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.WeaponEqpObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)

                            End If

                        End If

                        'desequipar casco
                        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.CascoEqpObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)

                            End If

                        End If

                        'desequipar herramienta
                        If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.AnilloEqpObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)

                            End If

                        End If

                        'desequipar municiones
                        If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.MunicionEqpObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)

                            End If

                        End If

                        'desequipar escudo
                        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.EscudoEqpObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)

                            End If

                        End If
                    
                        If UserList(UserIndex).Invent.MagicoObjIndex > 0 Then
                            If ItemSeCae(UserList(UserIndex).Invent.MagicoObjIndex) Then
                                Call Desequipar(UserIndex, UserList(UserIndex).Invent.MagicoSlot)

                            End If

                        End If

                    Else
                
                        Dim MiObj As obj

                        MiObj.Amount = 1
                        MiObj.ObjIndex = PENDIENTE
                        Call QuitarObjetos(PENDIENTE, 1, UserIndex)
                        Call MakeObj(MiObj, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
                        Call WriteConsoleMsg(UserIndex, "Has perdido tu pendiente del sacrificio.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)

                End If

            End If

        End If

    End If

    UserList(UserIndex).flags.CarroMineria = 0
    
    ' DESEQUIPA TODOS LOS OBJETOS
    
    If UserList(UserIndex).Char.Arma_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 1))
        UserList(UserIndex).Char.Arma_Aura = ""

    End If

    If UserList(UserIndex).Char.Body_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 2))
        UserList(UserIndex).Char.Body_Aura = 0

    End If
    
    If UserList(UserIndex).Char.Escudo_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 3))
        UserList(UserIndex).Char.Escudo_Aura = 0

    End If
    
    If UserList(UserIndex).Char.Head_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 4))
        UserList(UserIndex).Char.Head_Aura = 0

    End If

    If UserList(UserIndex).Char.Otra_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 5))
        UserList(UserIndex).Char.Otra_Aura = 0

    End If
                        
    'desequipar montura
    If UserList(UserIndex).flags.Montado > 0 Then

        Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

    End If
    
    ' If UserList(UserIndex).flags.Navegando > 0 Then

    '  Call DoNavega(UserIndex, ObjData(UserList(UserIndex).Invent.BarcoObjIndex), UserList(UserIndex).Invent.BarcoSlot)
    'End If
        
    UserList(UserIndex).Char.speeding = VelocidadMuerto
    'Call WriteVelocidadToggle(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(UserIndex).Char.loops = INFINITE_LOOPS Then
        UserList(UserIndex).Char.FX = 0
        UserList(UserIndex).Char.loops = 0

    End If
    
    UserList(UserIndex).flags.VecesQueMoriste = UserList(UserIndex).flags.VecesQueMoriste + 1
    
    ' << Restauramos los atributos >>
    If UserList(UserIndex).flags.TomoPocion = True And UserList(UserIndex).flags.BattleModo = 0 Then

        For i = 1 To 4
            UserList(UserIndex).Stats.UserAtributos(i) = UserList(UserIndex).Stats.UserAtributosBackUP(i)
        Next i

        Call WriteFYA(UserIndex)

    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(UserIndex).flags.Navegando = 0 Then
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
        
    Else

        If ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje = iTraje Then
            UserList(UserIndex).Char.Body = iRopaBuceoMuerto
            UserList(UserIndex).Char.Head = iCabezaMuerto
        Else
            UserList(UserIndex).Char.Body = iFragataFantasmal ';)
            UserList(UserIndex).Char.Head = 0

        End If

    End If
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, NingunArma, NingunEscudo, NingunCasco)
    Call WriteUpdateUserStats(UserIndex)
    
    'If UCase$(MapInfo(UserList(UserIndex).Pos.Map).restrict_mode) = "NEWBIE" Then
    '    UserList(UserIndex).flags.pregunta = 5
    '    Call WritePreguntaBox(UserIndex, "¡Has muerto! ¿Deseas ser resucitado?")
    'End If

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    If Abs(CInt(UserList(Muerto).Stats.ELV) - CInt(UserList(Atacante).Stats.ELV)) > 14 Then Exit Sub
    If Status(Muerto) = 0 Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name

            If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1

        End If
        
        If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho
            
            'con esto evitamos que se vuelva a reenlistar
        End If

    ElseIf Status(Muerto) = 1 Then

        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name

            If UserList(Atacante).Faccion.CiudadanosMatados < MAXUSERMATADOS Then UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1

        End If

    End If

End Sub

Sub ContarPuntoBattle(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If UserList(Muerto).flags.LevelBackup < 40 And UserList(Atacante).flags.LevelBackup < 40 Then Exit Sub
    If Abs(CInt(UserList(Muerto).flags.LevelBackup) - CInt(UserList(Atacante).flags.LevelBackup)) > 5 Then Exit Sub

    If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
        UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name
            
        UserList(Atacante).flags.BattlePuntos = UserList(Atacante).flags.BattlePuntos + 1
        UserList(Muerto).flags.BattlePuntos = UserList(Muerto).flags.BattlePuntos - 1
            
        Call WriteConsoleMsg(Atacante, "Has ganado 1 punto battle.", FontTypeNames.FONTTYPE_EXP)
        Call WriteConsoleMsg(Muerto, "Has perdido 1 punto battle.", FontTypeNames.FONTTYPE_EXP)
        Call CheckRanking(Battle, Atacante, UserList(Atacante).flags.BattlePuntos)
        Call CheckRanking(Battle, Muerto, UserList(Muerto).flags.BattlePuntos)
        Call GuardarRanking

    End If

End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef obj As obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
    '**************************************************************
    Dim Notfound As Boolean

    Dim LoopC    As Integer

    Dim tX       As Integer

    Dim tY       As Integer

    Dim hayobj   As Boolean

    hayobj = False
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.x, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do

        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.x - LoopC To Pos.x + LoopC
            
                If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the Amount dropped + Amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)

                    If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.Amount + obj.Amount > MAX_INVENTORY_OBJS)

                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.x = tX
                        nPos.Y = tY
                        tX = Pos.x + LoopC
                        tY = Pos.Y + LoopC

                    End If

                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.x = 0
        nPos.Y = 0

    End If

End Sub

Sub WarpToLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte, Optional ByVal FX As Boolean = False)
    'Santo: Sub para buscar la posición legal mas cercana al objetivo y warpearlo.

    Dim ALoop As Byte, Find As Boolean, lX As Long, lY As Long

    Find = False
    ALoop = 1

    Do Until Find = True

        For lX = x - ALoop To x + ALoop
            For lY = Y - ALoop To Y + ALoop

                With MapData(Map, lX, lY)

                    If .UserIndex <= 0 Then
                        If .Blocked <> 1 Then
                            If .TileExit.Map = 0 Then
                                If .NpcIndex <= 0 Then
                                    Call WarpUserChar(UserIndex, Map, lX, lY, FX)
                                    Find = True
                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                End With

            Next lY
        Next lX

        ALoop = ALoop + 1
    Loop

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

    Dim OldMap As Integer

    Dim OldX   As Integer

    Dim OldY   As Integer
    
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(UserIndex).ComUsu.DestUsu)

            End If

        End If

    End If
    
    'Quitar el dialogo
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
    Call WriteRemoveAllDialogs(UserIndex)
    
    OldMap = UserList(UserIndex).Pos.Map
    OldX = UserList(UserIndex).Pos.x
    OldY = UserList(UserIndex).Pos.Y
    
    Call EraseUserChar(UserIndex, True)
    
    If OldMap <> Map Then
        
        Call WriteChangeMap(UserIndex, Map)
        'Call WriteLight(UserIndex, map)
        Call WriteHora(UserIndex)
        
        ' If MapInfo(OldMap).music_numberLow <> MapInfo(map).music_numberLow Then
        'Call WritePlayMidi(UserIndex, MapInfo(map).music_numberLow, 1)
        'End If
        
        If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And UserList(UserIndex).Stats.ELV < 42 Then
            Call WriteConsoleMsg(UserIndex, "Estas saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.", FontTypeNames.FONTTYPE_WARNING)

        End If
        
        UserList(UserIndex).flags.NecesitaOxigeno = RequiereOxigeno(Map)

        If UserList(UserIndex).flags.NecesitaOxigeno Then
            Call WriteContadores(UserIndex)
            Call WriteOxigeno(UserIndex)

            If UserList(UserIndex).Counters.Oxigeno = 0 Then
                UserList(UserIndex).flags.Ahogandose = 1

            End If

        End If

        If RequiereOxigeno(OldMap) = True And UserList(UserIndex).flags.NecesitaOxigeno = False Then  'And UserList(UserIndex).Stats.ELV < 35 Then
        
            'Call WriteConsoleMsg(UserIndex, "Ya no necesitas oxigeno.", FontTypeNames.FONTTYPE_WARNING)
            Call WriteContadores(UserIndex)
            Call WriteOxigeno(UserIndex)

        End If
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0

        End If

    End If
    
    UserList(UserIndex).Pos.x = x
    UserList(UserIndex).Pos.Y = Y
    UserList(UserIndex).Pos.Map = Map
    
    If FX Then
        Call MakeUserChar(True, Map, UserIndex, Map, x, Y, 1)
    Else
        Call MakeUserChar(True, Map, UserIndex, Map, x, Y, 0)

    End If
    
    Call WriteUserCharIndexInServer(UserIndex)
    
    'Force a flush, so user index is in there before it's destroyed for teleporting
    Call FlushBuffer(UserIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

    End If
    
    'Reparacion temporal del bug de particulas. 08/07/09 LADDER

    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 71, False))
    
    If FX And UserList(UserIndex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, Y))
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.LogeoLevel1, 200, False))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, ParticulasIndex.LogeoLevel1, 80))

    End If
    
    If MapInfo(Map).zone = "DUNGEON" Then
        If UserList(UserIndex).flags.Montado > 0 Then
            Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

        End If

    End If
    
    ' Call WarpFamiliar(UserIndex)
End Sub

Sub WarpFamiliar(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If .Familiar.Invocado = 1 Then
            Call QuitarNPC(.Familiar.Id)
            ' If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
            .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, UserList(UserIndex).Pos, False, True)

            'Controlamos que se sumoneo OK
            If .Familiar.Id = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota. Se provoco un ERROR.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            Call CargarFamiliar(UserIndex)
        Else

            'Call WriteConsoleMsg(UserIndex, "No se permiten familiares en zona segura. " & .Familiar.Nombre & " te esperará afuera.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    End With
            
End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
 
    If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
        UserList(UserIndex).Counters.Saliendo = True

        If UserList(UserIndex).flags.Privilegios = PlayerType.user And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 And UserList(UserIndex).flags.Muerto = 0 Then
            UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
            Call WriteLocaleMsg(UserIndex, "203", FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Counters.Salir)
            'Call WriteConsoleMsg(UserIndex, "Saliendo...Se saldrá del juego en " & UserList(UserIndex).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        Else
            
            'Call WriteConsoleMsg(UserIndex, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
            Call WriteDisconnect(UserIndex)
            Call FlushBuffer(UserIndex)
  
            Call CloseSocket(UserIndex)

        End If

    End If

End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/08
    '
    '***************************************************
    If UserList(UserIndex).Counters.Saliendo Then

        ' Is the user still connected?
        If UserList(UserIndex).ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else

            'Simply reset
            If UserList(UserIndex).flags.Privilegios = PlayerType.user And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 Then
                UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
            Else
                Call WriteConsoleMsg(UserIndex, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
                Call WriteDisconnect(UserIndex)
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)

            End If
            
            'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
        End If

    End If

End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)

    Dim ViejoNick       As String

    Dim ViejoCharBackup As String

    If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
    ViejoNick = UserList(UserIndexDestino).name

    If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
        'hace un backup del char
        ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
        Name CharPath & ViejoNick & ".chr" As ViejoCharBackup

    End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)

    If FileExist(CharPath & nombre & ".chr", vbArchive) = False Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Vitalidad: " & GetVar(CharPath & nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
    
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
    
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Veces Que Murio: " & GetVar(CharPath & nombre & ".chr", "Flags", "VecesQueMoriste"), FontTypeNames.FONTTYPE_INFO)
        #If ConUpTime Then

            Dim TempSecs As Long

            Dim TempStr  As String

            TempSecs = GetVar(CharPath & nombre & ".chr", "INIT", "UpTime")
            TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If

    End If

End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    On Error Resume Next

    Dim j        As Integer

    Dim CharFile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long

    CharFile = CharPath & CharName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente
    '**************************************************************
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

    If UserList(UserIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
   
        If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)

    End If

    If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Exit Sub

    UserList(UserIndex).Faccion.Status = 0

    Call RefreshCharStatus(UserIndex)

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente.
    '**************************************************************

    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

    UserList(UserIndex).Faccion.Status = 1
    Call RefreshCharStatus(UserIndex)

End Sub

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
    '***************************************************
    'Author: Unknown
    'Last Modification: 30/09/2020
    '
    '***************************************************

    If UserList(UserIndex).Stats.InventLevel > 0 Then
        getMaxInventorySlots = MAX_USERINVENTORY_SLOTS + UserList(UserIndex).Stats.InventLevel * SLOTS_PER_ROW_INVENTORY
    Else
        getMaxInventorySlots = MAX_USERINVENTORY_SLOTS

    End If

End Function
