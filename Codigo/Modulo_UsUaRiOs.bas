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
        
        On Error GoTo ActStats_Err
        

        Dim DaExp       As Integer

        Dim EraCriminal As Byte
    
100     DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)
    
102     If UserList(attackerIndex).Stats.ELV < STAT_MAXELV Then
104         UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp

106         If UserList(attackerIndex).Stats.Exp > MAXEXP Then UserList(attackerIndex).Stats.Exp = MAXEXP

108         Call WriteUpdateExp(attackerIndex)
110         Call CheckUserLevel(attackerIndex)

        End If
    
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
    
112     Call WriteLocaleMsg(attackerIndex, "184", FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).name)
114     Call WriteLocaleMsg(attackerIndex, "140", FontTypeNames.FONTTYPE_EXP, DaExp)
          
        'Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
116     Call WriteLocaleMsg(VictimIndex, "185", FontTypeNames.FONTTYPE_FIGHT, UserList(attackerIndex).name)
    
118     If TriggerZonaPelea(VictimIndex, attackerIndex) <> TRIGGER6_PERMITE Then
120         EraCriminal = Status(attackerIndex)
        
122         If EraCriminal = 2 And Status(attackerIndex) < 2 Then
124             Call RefreshCharStatus(attackerIndex)
126         ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
128             Call RefreshCharStatus(attackerIndex)

            End If

        End If
    
130     Call UserDie(VictimIndex)
    
132     If UserList(attackerIndex).flags.BattleModo = 1 Then
134         Call ContarPuntoBattle(VictimIndex, attackerIndex)

        End If
    
136     If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1
        'Call CheckearRecompesas(attackerIndex, 2)
    
    

        
        Exit Sub

ActStats_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ActStats", Erl)
        Resume Next
        
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer)
        
        On Error GoTo RevivirUsuario_Err
        

100     UserList(UserIndex).Char.speeding = VelocidadNormal
        'Call WriteVelocidadToggle(UserIndex)
102     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))

104     UserList(UserIndex).flags.Muerto = 0
106     UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)

108     If UserList(UserIndex).Stats.MinHp <> UserList(UserIndex).Stats.MaxHp Then
110         UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp

        End If

112     If UserList(UserIndex).flags.Navegando = 1 Then

            Dim Barco As ObjData

114         Barco = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)

116         If Barco.Ropaje <> iTraje Then
118             UserList(UserIndex).Char.Head = 0
120             UserList(UserIndex).Char.CascoAnim = NingunCasco

            End If
    
122         If Barco.Ropaje = iTraje Then
124             UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head
126             Call WriteNadarToggle(UserIndex, True)
        
            Else
128             Call WriteNadarToggle(UserIndex, False)
        
            End If
    
130         If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
132             If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
134             If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaCiuda
136             If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraCiuda
138             If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonCiuda
140         ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then

142             If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
144             If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarcaPk
146             If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGaleraPk
148             If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleonPk
            Else

150             If Barco.Ropaje = iTraje Then UserList(UserIndex).Char.Body = iTraje
152             If Barco.Ropaje = iBarca Then UserList(UserIndex).Char.Body = iBarca
154             If Barco.Ropaje = iGalera Then UserList(UserIndex).Char.Body = iGalera
156             If Barco.Ropaje = iGaleon Then UserList(UserIndex).Char.Body = iGaleon

            End If
    
158         UserList(UserIndex).Char.ShieldAnim = NingunEscudo
160         UserList(UserIndex).Char.WeaponAnim = NingunArma
            ' UserList(UserIndex).Char.CascoAnim = NingunCasco
   
162         UserList(UserIndex).Char.speeding = Barco.Velocidad
            'Call WriteVelocidadToggle(UserIndex)
164         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.speeding))
   
        Else
    
176         UserList(UserIndex).Char.Head = UserList(UserIndex).OrigChar.Head

178         If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
180             UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

            End If

182         If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
184             UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

            End If

186         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
188             UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
    
190             If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CreaGRH <> "" Then
192                 UserList(UserIndex).Char.Arma_Aura = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).CreaGRH
194                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 1))

                End If
        
            End If

196         If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
198             UserList(UserIndex).Char.Body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
    
204             If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).CreaGRH <> "" Then
206                 UserList(UserIndex).Char.Body_Aura = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).CreaGRH
208                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Body_Aura, False, 2))

                End If

            Else
174             Call DarCuerpoDesnudo(UserIndex)
        
            End If

210         If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
212             UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim

214             If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).CreaGRH <> "" Then
216                 UserList(UserIndex).Char.Escudo_Aura = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).CreaGRH
218                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Escudo_Aura, False, 3))

                End If
        
            End If

220         If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
222             UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim

224             If ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CreaGRH <> "" Then
226                 UserList(UserIndex).Char.Head_Aura = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CreaGRH
228                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Head_Aura, False, 4))

                End If
        
            End If

230         If UserList(UserIndex).Invent.MagicoObjIndex > 0 Then
232             If ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CreaGRH <> "" Then
234                 UserList(UserIndex).Char.Otra_Aura = ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CreaGRH
236                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Otra_Aura, False, 5))

                End If

            End If

238         If UserList(UserIndex).Invent.NudilloObjIndex > 0 Then
240             If ObjData(UserList(UserIndex).Invent.NudilloObjIndex).CreaGRH <> "" Then
242                 UserList(UserIndex).Char.Arma_Aura = ObjData(UserList(UserIndex).Invent.NudilloObjIndex).CreaGRH
244                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 1))

                End If
            End If
            
            If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                If ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).CreaGRH <> "" Then
                    UserList(UserIndex).Char.Anillo_Aura = ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Anillo_Aura, False, 6))
                End If
            End If

        End If

246     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

248     Call WriteUpdateUserStats(UserIndex)

250     'Call WriteHora(UserIndex)

        
        Exit Sub

RevivirUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.RevivirUsuario", Erl)
        Resume Next
        
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
        
        On Error GoTo ChangeUserChar_Err
        

100     With UserList(UserIndex).Char
102         .Body = Body
104         .Head = Head
106         .Heading = Heading
108         .WeaponAnim = Arma
110         .ShieldAnim = Escudo
112         .CascoAnim = Casco

        End With
    
114     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, Head, Heading, UserList(UserIndex).Char.CharIndex, Arma, Escudo, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, Casco))

        
        Exit Sub

ChangeUserChar_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ChangeUserChar", Erl)
        Resume Next
        
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
    
    MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
    Error = "5"
    UserList(UserIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    Error = "6"
    Exit Sub
    
ErrorHandler:
    Call LogError("Error en EraseUserchar " & Error & " - " & Err.Number & ": " & Err.description)

End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
        
        On Error GoTo RefreshCharStatus_Err
        

        '*************************************************
        'Author: Tararira
        'Last modified: 6/04/2007
        'Refreshes the status and tag of UserIndex.
        '*************************************************
        Dim klan As String

100     If UserList(UserIndex).GuildIndex > 0 Then
102         klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
104         klan = " <" & klan & ">"

        End If
    
106     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.Status, UserList(UserIndex).name & klan))

        
        Exit Sub

RefreshCharStatus_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.RefreshCharStatus", Erl)
        Resume Next
        
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal appear As Byte = 0)

    On Error GoTo hayerror

    Dim CharIndex As Integer

    Dim errort    As String

    If InMapBounds(Map, X, Y) Then

        'If needed make a new character in list
        If UserList(UserIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(UserIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = UserIndex
        End If

        errort = "1"
        
        'Place character on map if needed
        If toMap Then MapData(Map, X, Y).UserIndex = UserIndex
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
                Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name & " <" & klan & ">", bCr, UserList(UserIndex).flags.Privilegios, UserList(UserIndex).Char.ParticulaFx, UserList(UserIndex).Char.Head_Aura, UserList(UserIndex).Char.Arma_Aura, UserList(UserIndex).Char.Body_Aura, UserList(UserIndex).Char.Anillo_Aura, UserList(UserIndex).Char.Otra_Aura, UserList(UserIndex).Char.Escudo_Aura, UserList(UserIndex).Char.speeding, False, UserList(UserIndex).donador.activo, appear, UserList(UserIndex).Grupo.Lider, UserList(UserIndex).GuildIndex, clan_nivel, UserList(UserIndex).Stats.MinHp, UserList(UserIndex).Stats.MaxHp, 0)
            Else
                errort = "6"
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map, appear)
            End If

        Else 'if tiene clan

            If Not toMap Then
                errort = "7"
                Call WriteCharacterCreate(sndIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.CharIndex, X, Y, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.FX, 999, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).name, bCr, UserList(UserIndex).flags.Privilegios, UserList(UserIndex).Char.ParticulaFx, UserList(UserIndex).Char.Head_Aura, UserList(UserIndex).Char.Arma_Aura, UserList(UserIndex).Char.Body_Aura, UserList(UserIndex).Char.Anillo_Aura, UserList(UserIndex).Char.Otra_Aura, UserList(UserIndex).Char.Escudo_Aura, UserList(UserIndex).Char.speeding, False, UserList(UserIndex).donador.activo, appear, UserList(UserIndex).Grupo.Lider, 0, 0, UserList(UserIndex).Stats.MinHp, UserList(UserIndex).Stats.MaxHp, 0)
            Else
                errort = "8"
                Call AgregarUser(UserIndex, UserList(UserIndex).Pos.Map, appear)
            End If

        End If 'if clan

    End If

    Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description & " - Nombre del usuario " & UserList(UserIndex).name) & " - " & errort & "- Pos: M: " & Map & " X: " & X & " Y: " & Y
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

    On Error GoTo ErrHandler

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

        WasNewbie = EsNewbie(UserIndex)
        
        Do While .Stats.Exp >= .Stats.ELU And .Stats.ELV < STAT_MAXELV
            
            'Store it!
            'Call Statistics.UserLevelUp(UserIndex)

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 106, 0))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
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
            
            If OroPorNivel > 0 Then
                If EsNewbie(UserIndex) Then
                    Dim OroRecompenza As Long
    
                    OroRecompenza = OroPorNivel * .Stats.ELV * OroMult * .flags.ScrollOro
                    .Stats.GLD = .Stats.GLD + OroRecompenza
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & OroRecompenza & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, "29", FontTypeNames.FONTTYPE_INFO, OroRecompenza)
                End If
            End If
            
            If Not EsNewbie(UserIndex) And WasNewbie Then
        
                Call QuitarNewbieObj(UserIndex)
            
            End If
        
        Loop
        
        If PasoDeNivel Then
            'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
            End If
        
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

ErrHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub

Function PuedeAtravesarAgua(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo PuedeAtravesarAgua_Err
        

100     PuedeAtravesarAgua = UserList(UserIndex).flags.Navegando = 1
        'If PuedeAtravesarAgua = True Then
        '  Exit Function
        'Else
        '  If UserList(UserIndex).flags.Nadando = 1 Then
        ' PuedeAtravesarAgua = True
        'End If
        'End If

        
        Exit Function

PuedeAtravesarAgua_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.PuedeAtravesarAgua", Erl)
        Resume Next
        
End Function

Sub MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading)
        
        On Error GoTo MoveUserChar_Err
        

        Dim nPos         As WorldPos

        Dim nPosOriginal As WorldPos

        Dim nPosMuerto   As WorldPos

        Dim sailing      As Boolean

100     With UserList(UserIndex)

102         If .Accion.AccionPendiente = True Then
104             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, .Accion.Particula, 1, True))
106             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 1, Accion_Barra.CancelarAccion))
108             .Accion.AccionPendiente = False
110             .Accion.Particula = 0
112             .Accion.TipoAccion = Accion_Barra.CancelarAccion
114             .Accion.HechizoPendiente = 0
116             .Accion.RunaObj = 0
118             .Accion.ObjSlot = 0
120             .Accion.AccionPendiente = False

            End If

122         sailing = PuedeAtravesarAgua(UserIndex)
124         nPos = .Pos
126         Call HeadtoPos(nHeading, nPos)
        
128         If MapData(nPos.Map, nPos.X, nPos.Y).TileExit.Map <> 0 And .Counters.TiempoDeMapeo > 0 Then
130             If .flags.Muerto = 0 Then
132                 Call WriteConsoleMsg(UserIndex, "Estas en combate, debes aguardar " & .Counters.TiempoDeMapeo & " segundo(s) para escapar...", FontTypeNames.FONTTYPE_INFOBOLD)
134                 Call WritePosUpdate(UserIndex)
                    Exit Sub
    
                End If
    
            End If
    
136         If MapData(nPos.Map, nPos.X, nPos.Y).UserIndex <> 0 Then
                Dim IndexMuerto As Integer
                IndexMuerto = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                If UserList(IndexMuerto).flags.Muerto = 1 Or UserList(IndexMuerto).flags.AdminInvisible = 1 Then

142                 Call WarpToLegalPos(IndexMuerto, UserList(IndexMuerto).Pos.Map, UserList(IndexMuerto).Pos.X, UserList(IndexMuerto).Pos.Y, False)
    
                Else
166                 Call WritePosUpdate(UserIndex)
    
                    'Call WritePosUpdate(MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
                End If
    
            End If
    
168         If LegalWalk(.Pos.Map, nPos.X, nPos.Y, nHeading, sailing, Not sailing, .flags.Montado) Then
170             If MapInfo(.Pos.Map).NumUsers > 1 Then
                    'si no estoy solo en el mapa...
    
172                 Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
            
                End If
    
                'Call RefreshAllUser(UserIndex) '¿Clones? Ladder probar
                'Update map and user pos
174             MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
176             .Pos = nPos
178             .Char.Heading = nHeading
180             MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
            
                'Actualizamos las áreas de ser necesario
182             Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)
           
            Else
184             Call WritePosUpdate(UserIndex)
    
            End If
        
186         If .Counters.Trabajando Then
188             Call WriteMacroTrabajoToggle(UserIndex, False)
    
            End If
    
190         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

        End With

        
        Exit Sub

MoveUserChar_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.MoveUserChar", Erl)
        Resume Next
        
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

    '*************************************************
    'Author: ZaMa
    'Last modified: 30/03/2009
    'Returns the heading opposite to the one passed by val.
    '*************************************************
    Select Case nHeading

        Case eHeading.EAST
            InvertHeading = eHeading.WEST

        Case eHeading.WEST
            InvertHeading = eHeading.EAST

        Case eHeading.SOUTH
            InvertHeading = eHeading.NORTH

        Case eHeading.NORTH
            InvertHeading = eHeading.SOUTH

    End Select

End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal slot As Byte, ByRef Object As UserOBJ)
        
        On Error GoTo ChangeUserInv_Err
        
100     UserList(UserIndex).Invent.Object(slot) = Object
102     Call WriteChangeInventorySlot(UserIndex, slot)

        
        Exit Sub

ChangeUserInv_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ChangeUserInv", Erl)
        Resume Next
        
End Sub

Function NextOpenCharIndex() As Integer
        
        On Error GoTo NextOpenCharIndex_Err
        

        Dim LoopC As Long
    
100     For LoopC = 1 To MAXCHARS

102         If CharList(LoopC) = 0 Then
104             NextOpenCharIndex = LoopC
106             NumChars = NumChars + 1
            
108             If LoopC > LastChar Then LastChar = LoopC
            
                Exit Function

            End If

110     Next LoopC

        
        Exit Function

NextOpenCharIndex_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.NextOpenCharIndex", Erl)
        Resume Next
        
End Function

Function NextOpenUser() As Integer
        
        On Error GoTo NextOpenUser_Err
        

        Dim LoopC As Long
   
100     For LoopC = 1 To MaxUsers + 1

102         If LoopC > MaxUsers Then Exit For
104         If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
106     Next LoopC
   
108     NextOpenUser = LoopC

        
        Exit Function

NextOpenUser_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.NextOpenUser", Erl)
        Resume Next
        
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserStatsTxt_Err
        

        Dim GuildI As Integer

100     Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & UserList(UserIndex).Stats.ELU, FontTypeNames.FONTTYPE_INFO)
104     Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHp & "/" & UserList(UserIndex).Stats.MaxHp & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
106     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
108         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit, FontTypeNames.FONTTYPE_INFO)

        End If
    
112     If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
114         If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
118             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
120         Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
    
122     If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
124         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
126         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
    
128     GuildI = UserList(UserIndex).GuildIndex

130     If GuildI > 0 Then
132         Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)

134         If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
136             Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)

            End If

            'guildpts no tienen objeto
        End If
    
        #If ConUpTime Then

            Dim TempDate As Date

            Dim TempSecs As Long

            Dim TempStr  As String

138         TempDate = Now - UserList(UserIndex).LogOnTime
140         TempSecs = (UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
142         TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
144         Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
146         Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If

148     Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map, FontTypeNames.FONTTYPE_INFO)
150     Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma), FontTypeNames.FONTTYPE_INFO)
152     Call WriteConsoleMsg(sendIndex, "Veces que Moriste: " & UserList(UserIndex).flags.VecesQueMoriste, FontTypeNames.FONTTYPE_INFO)

        Exit Sub

SendUserStatsTxt_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserStatsTxt", Erl)
        Resume Next
        
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserMiniStatsTxt_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is online.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
100     With UserList(UserIndex)
102         Call WriteConsoleMsg(sendIndex, "Pj: " & .name, FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(sendIndex, "Ciudadanos Matados: " & .Faccion.CiudadanosMatados & " Criminales Matados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)

112         If .GuildIndex > 0 Then
114             Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)

            End If

116         Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & .Stats.GLD, FontTypeNames.FONTTYPE_INFO)
118         Call WriteConsoleMsg(sendIndex, "Oro en banco: " & .Stats.Banco, FontTypeNames.FONTTYPE_INFO)
    
120         Call WriteConsoleMsg(sendIndex, "Cuenta: " & .Cuenta, FontTypeNames.FONTTYPE_INFO)
122         Call WriteConsoleMsg(sendIndex, "Creditos: " & .donador.CreditoDonador, FontTypeNames.FONTTYPE_INFO)
124         Call WriteConsoleMsg(sendIndex, "Fecha Vencimiento Donador: " & .donador.FechaExpiracion, FontTypeNames.FONTTYPE_INFO)
    
        End With

        
        Exit Sub

SendUserMiniStatsTxt_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)
        Resume Next
        
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
        
        On Error GoTo SendUserMiniStatsTxtFromChar_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is offline.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
        Dim CharFile      As String

        Dim Ban           As String

        Dim BanDetailPath As String

100     BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
102     CharFile = CharPath & CharName & ".chr"
    
104     If FileExist(CharFile) Then
106         Call WriteConsoleMsg(sendIndex, "Pj: " & CharName, FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
112         Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
114         Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
116         Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & GetVar(CharFile, "STATS", "GLD"), FontTypeNames.FONTTYPE_INFO)
118         Call WriteConsoleMsg(sendIndex, "Oro en boveda: " & GetVar(CharFile, "STATS", "BANCO"), FontTypeNames.FONTTYPE_INFO)
120         Call WriteConsoleMsg(sendIndex, "Cuenta: " & GetVar(CharFile, "INIT", "Cuenta"), FontTypeNames.FONTTYPE_INFO)
        
122         If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
124             Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)

            End If
        
126         Ban = GetVar(CharFile, "BAN", "BanMotivo")
128         Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)

130         If Ban = "1" Then
132             Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason"), FontTypeNames.FONTTYPE_INFO)

            End If
        
        Else
134         Call WriteConsoleMsg(sendIndex, "El pj no existe: " & CharName, FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

SendUserMiniStatsTxtFromChar_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserMiniStatsTxtFromChar", Erl)
        Resume Next
        
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
        
        On Error GoTo DameUserIndex_Err
        

        Dim LoopC As Integer
  
100     LoopC = 1
  
102     Do Until UserList(LoopC).ConnID = SocketId

104         LoopC = LoopC + 1
    
106         If LoopC > MaxUsers Then
108             DameUserIndex = 0
                Exit Function

            End If
    
        Loop
  
110     DameUserIndex = LoopC

        
        Exit Function

DameUserIndex_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.DameUserIndex", Erl)
        Resume Next
        
End Function

Function DameUserIndexConNombre(ByVal nombre As String) As Integer
        
        On Error GoTo DameUserIndexConNombre_Err
        

        Dim LoopC As Integer
  
100     LoopC = 1
  
102     nombre = UCase$(nombre)

104     Do Until UCase$(UserList(LoopC).name) = nombre

106         LoopC = LoopC + 1
    
108         If LoopC > MaxUsers Then
110             DameUserIndexConNombre = 0
                Exit Function

            End If
    
        Loop
  
112     DameUserIndexConNombre = LoopC

        
        Exit Function

DameUserIndexConNombre_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.DameUserIndexConNombre", Erl)
        Resume Next
        
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo NPCAtacado_Err
        

        '**********************************************
        'Author: Unknown
        'Last Modification: 24/07/2007
        '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
        '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
        '**********************************************
        Dim EraCriminal As Byte

        'Guardamos el usuario que ataco el npc.
100     If Npclist(NpcIndex).Movement <> ESTATICO And Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
102         Npclist(NpcIndex).Target = UserIndex
104         Npclist(NpcIndex).Movement = TipoAI.NpcMaloAtacaUsersBuenos
106         Npclist(NpcIndex).flags.AttackedBy = UserList(UserIndex).name
        End If

        'Npc que estabas atacando.
        Dim LastNpcHit As Integer

108     LastNpcHit = UserList(UserIndex).flags.NPCAtacado
        'Guarda el NPC que estas atacando ahora.
110     UserList(UserIndex).flags.NPCAtacado = NpcIndex

        'Revisamos robo de npc.
        'Guarda el primer nick que lo ataca.
112     If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then

            'El que le pegabas antes ya no es tuyo
114         If LastNpcHit <> 0 Then
116             If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).name Then
118                 Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

                End If

            End If

120         Npclist(NpcIndex).flags.AttackedFirstBy = UserList(UserIndex).name
122     ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(UserIndex).name Then

            'Estas robando NPC
            'El que le pegabas antes ya no es tuyo
124         If LastNpcHit <> 0 Then
126             If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(UserIndex).name Then
128                 Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

                End If

            End If

        End If

        '  EraCriminal = Status(UserIndex)

130     If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
132         If Status(UserIndex) = 1 Or Status(UserIndex) = 3 Then
134             Call VolverCriminal(UserIndex)

            End If

        End If
        
        If Npclist(NpcIndex).MaestroUser > 0 Then
            If Npclist(NpcIndex).MaestroUser <> UserIndex Then
                Call AllMascotasAtacanUser(UserIndex, Npclist(NpcIndex).MaestroUser)
            End If
        End If

        Call CheckPets(NpcIndex, UserIndex, False)
        
        Exit Sub

NPCAtacado_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.NPCAtacado", Erl)
        Resume Next
        
End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo PuedeApuñalar_Err
        

100     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
102         PuedeApuñalar = ((UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UserList(UserIndex).clase = eClass.Assasin) And (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
        Else
104         PuedeApuñalar = False

        End If

        
        Exit Function

PuedeApuñalar_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.PuedeApuñalar", Erl)
        Resume Next
        
End Function

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
        
        On Error GoTo SubirSkill_Err
        

100     If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

102     If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
        
            Dim Lvl As Integer

104         Lvl = UserList(UserIndex).Stats.ELV
        
106         If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
            
108         If UserList(UserIndex).Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
        
            Dim Aumenta As Integer

            Dim Prob    As Integer

            Dim Menor   As Byte

110         Menor = 10
             
112         If Lvl <= 3 Then
114             Prob = 25
116         ElseIf Lvl > 3 And Lvl < 6 Then
118             Prob = 27
120         ElseIf Lvl >= 6 And Lvl < 10 Then
122             Prob = 30
124         ElseIf Lvl >= 10 And Lvl < 20 Then
126             Prob = 33
            Else
128             Prob = 38
            End If
             
130         Aumenta = RandomNumber(1, Prob * DificultadSubirSkill)
             
132         If UserList(UserIndex).flags.PendienteDelExperto = 1 Then
134             Menor = 15

            End If
        
136         If Aumenta < Menor Then
138             UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
    
140             Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
            
                Dim BonusExp As Long

142             BonusExp = 5 * ExpMult * UserList(UserIndex).flags.ScrollExp
        
144             If UserList(UserIndex).donador.activo = 1 Then
146                 BonusExp = BonusExp * 1.1

                End If
        
148             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
150                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + BonusExp

152                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
            
154                 If UserList(UserIndex).ChatCombate = 1 Then
156                     Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, BonusExp)

                    End If
                
158                 Call WriteUpdateExp(UserIndex)
160                 Call CheckUserLevel(UserIndex)

                End If

            End If

        End If

        
        Exit Sub

SubirSkill_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SubirSkill", Erl)
        Resume Next
        
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
    
    With UserList(UserIndex)
    
        'Sonido
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MUERTE_HOMBRE, .Pos.X, .Pos.Y))
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Ahogandose = 0
        .flags.Incinerado = 0
        .flags.incinera = 0
        .flags.Paraliza = 0
        .flags.Envenena = 0
        .flags.Estupidiza = 0
        .flags.Muerto = 1
        '.flags.SeguroParty = True
        'Call WritePartySafeOn(UserIndex)
        
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MUERTE_HOMBRE, .Pos.X, .Pos.Y))
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
    
    .Stats.MinHp = 0
    .Stats.MinSta = 0
    .flags.AtacadoPorUser = 0
    .flags.Envenenado = 0
    .flags.Ahogandose = 0
    .flags.Incinerado = 0
    .flags.incinera = 0
    .flags.Paraliza = 0
    .flags.Envenena = 0
    .flags.Estupidiza = 0
    .flags.Muerto = 1
    '.flags.SeguroParty = True
    'Call WritePartySafeOn(UserIndex)
    
    aN = .flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString

    End If
    
    aN = .flags.NPCAtacado

    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = .name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString

        End If

    End If

    .flags.AtacadoPorNpc = 0
    .flags.NPCAtacado = 0
    
    '<<<< Paralisis >>>>
    If .flags.Paralizado = 1 Then
        .flags.Paralizado = 0
        Call WriteParalizeOK(UserIndex)

    End If
    
    '<<<< Inmovilizado >>>>
    If .flags.Inmovilizado = 1 Then
        .flags.Inmovilizado = 0
        Call WriteInmovilizaOK(UserIndex)

    End If
    
    '<<< Estupidez >>>
    If .flags.Estupidez = 1 Then
        .flags.Estupidez = 0
        Call WriteDumbNoMore(UserIndex)

    End If
    
    '<<<< Descansando >>>>
    If .flags.Descansar Then
        .flags.Descansar = False
        Call WriteRestOK(UserIndex)

    End If
    
    '<<<< Meditando >>>>
    If .flags.Meditando Then
        .flags.Meditando = False
        .Char.FX = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
    End If
    
    'If .Familiar.Invocado = 1 Then
    ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso("17", Npclist(.Familiar.Id).Pos.x, Npclist(.Familiar.Id).Pos.Y))
    ' .Familiar.Invocado = 0
    ' Call QuitarNPC(.Familiar.Id)
    ' End If
    
    '<<<< Invisible >>>>
    If .flags.invisible = 1 Or .flags.Oculto = 1 Then
        .flags.Oculto = 0
        .flags.invisible = 0
        .Counters.TiempoOculto = 0
        .Counters.Invisibilidad = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

    End If

    If MapInfo(.Pos.Map).Seguro = 0 Then '  Ladder 06/07/2014 Si el mapa es seguro, no se caen los items
        If TriggerZonaPelea(UserIndex, UserIndex) <> eTrigger6.TRIGGER6_PERMITE Then

            ' << Si es newbie no pierde el inventario >>
            If (.flags.Privilegios And PlayerType.user) <> 0 Then
        
                If Not EsNewbie(UserIndex) Then
                    If .flags.PendienteDelSacrificio = 0 Then
                
                        Call TirarTodo(UserIndex)
                    
                        If .Invent.ArmourEqpObjIndex > 0 Then
                            If ItemSeCae(.Invent.ArmourEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                            End If

                        End If

                        'desequipar arma
                        If .Invent.WeaponEqpObjIndex > 0 Then
                            If ItemSeCae(.Invent.WeaponEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                            End If

                        End If

                        'desequipar casco
                        If .Invent.CascoEqpObjIndex > 0 Then
                            If ItemSeCae(.Invent.CascoEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.CascoEqpSlot)

                            End If

                        End If

                        'desequipar herramienta
                        If .Invent.AnilloEqpObjIndex > 0 Then
                            If ItemSeCae(.Invent.AnilloEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)

                            End If

                        End If

                        'desequipar municiones
                        If .Invent.MunicionEqpObjIndex > 0 Then
                            If ItemSeCae(.Invent.MunicionEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)

                            End If

                        End If

                        'desequipar escudo
                        If .Invent.EscudoEqpObjIndex > 0 Then
                            If ItemSeCae(.Invent.EscudoEqpObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)

                            End If

                        End If
                    
                        If .Invent.MagicoObjIndex > 0 Then
                            If ItemSeCae(.Invent.MagicoObjIndex) Then
                                Call Desequipar(UserIndex, .Invent.MagicoSlot)

                            End If
                        End If
    
                    Else
                
                        Dim MiObj As obj

                        MiObj.Amount = 1
                        MiObj.ObjIndex = PENDIENTE
                        Call QuitarObjetos(PENDIENTE, 1, UserIndex)
                        Call MakeObj(MiObj, .Pos.Map, .Pos.X, .Pos.Y)
                        Call WriteConsoleMsg(UserIndex, "Has perdido tu pendiente del sacrificio.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If EsNewbie(UserIndex) Then Call TirarTodosLosItemsNoNewbies(UserIndex)

                End If
    
            End If
    
        End If

    End If

    .flags.CarroMineria = 0
    
        .flags.CarroMineria = 0
        
        ' DESEQUIPA TODOS LOS OBJETOS
        
        If .Char.Arma_Aura <> "" Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 1))
            .Char.Arma_Aura = ""
    
    If .Char.Arma_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 1))
        .Char.Arma_Aura = ""

    End If

    If .Char.Body_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 2))
        .Char.Body_Aura = 0

    End If
    
    If .Char.Escudo_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 3))
        .Char.Escudo_Aura = 0

    End If
    
    If .Char.Head_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 4))
        .Char.Head_Aura = 0

    End If
    
    If .Char.Anillo_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 6))
        .Char.Anillo_Aura = 0
    End If

    If .Char.Otra_Aura <> "" Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 5))
        .Char.Otra_Aura = 0
    End If
                        
    'desequipar montura
    If .flags.Montado > 0 Then

        Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)

    End If
    
        End If
        
    .Char.speeding = VelocidadMuerto
    'Call WriteVelocidadToggle(UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If .Char.loops = INFINITE_LOOPS Then
        .Char.FX = 0
        .Char.loops = 0

    End If
    
    .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
    
    ' << Restauramos los atributos >>
    If .flags.TomoPocion = True And .flags.BattleModo = 0 Then

        For i = 1 To 4
            .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
        Next i

        Call WriteFYA(UserIndex)

    End If
    
    '<< Cambiamos la apariencia del char >>
    If .flags.Navegando = 0 Then
        .Char.Body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
        
    Else

        If ObjData(.Invent.BarcoObjIndex).Ropaje = iTraje Then
            .Char.Body = iRopaBuceoMuerto
            .Char.Head = iCabezaMuerto
        Else
            .Char.Body = iFragataFantasmal ';)
            .Char.Head = 0

        End If
        
        .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True And .flags.BattleModo = 0 Then
    
            For i = 1 To 4
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
    
            Call WriteFYA(UserIndex)
    
        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.Body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            
        Else
    
            If ObjData(.Invent.BarcoObjIndex).Ropaje = iTraje Then
                .Char.Body = iRopaBuceoMuerto
                .Char.Head = iCabezaMuerto
            Else
                .Char.Body = iFragataFantasmal ';)
                .Char.Head = 0
            End If
        End If
    
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    
        End If
        
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
            ' Si estan en agua o zona segura
            Else
                .MascotasType(i) = 0
            End If
        Next i
        
        .NroMascotas = 0
        
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
        'Call WriteUpdateUserStats(UserIndex)
        
        'If UCase$(MapInfo(.Pos.Map).restrict_mode) = "NEWBIE" Then
        '    .flags.pregunta = 5
        '    Call WritePreguntaBox(UserIndex, "¡Has muerto! ¿Deseas ser resucitado?")
        'End If
        
    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
        
        On Error GoTo ContarMuerte_Err
        

100     If EsNewbie(Muerto) Then Exit Sub
102     If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
104     If Abs(CInt(UserList(Muerto).Stats.ELV) - CInt(UserList(Atacante).Stats.ELV)) > 14 Then Exit Sub
106     If Status(Muerto) = 0 Then
108         If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
110             UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name

112             If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1

            End If
        
114         If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
116             UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho
            
                'con esto evitamos que se vuelva a reenlistar
            End If

118     ElseIf Status(Muerto) = 1 Then

120         If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
122             UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name

124             If UserList(Atacante).Faccion.CiudadanosMatados < MAXUSERMATADOS Then UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1

            End If

        End If

        
        Exit Sub

ContarMuerte_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ContarMuerte", Erl)
        Resume Next
        
End Sub

Sub ContarPuntoBattle(ByVal Muerto As Integer, ByVal Atacante As Integer)
        
        On Error GoTo ContarPuntoBattle_Err
        

100     If UserList(Muerto).flags.LevelBackup < 40 And UserList(Atacante).flags.LevelBackup < 40 Then Exit Sub
102     If Abs(CInt(UserList(Muerto).flags.LevelBackup) - CInt(UserList(Atacante).flags.LevelBackup)) > 5 Then Exit Sub

104     If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
106         UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name
            
108         UserList(Atacante).flags.BattlePuntos = UserList(Atacante).flags.BattlePuntos + 1
110         UserList(Muerto).flags.BattlePuntos = UserList(Muerto).flags.BattlePuntos - 1
            
112         Call WriteConsoleMsg(Atacante, "Has ganado 1 punto battle.", FontTypeNames.FONTTYPE_EXP)
114         Call WriteConsoleMsg(Muerto, "Has perdido 1 punto battle.", FontTypeNames.FONTTYPE_EXP)
116         Call CheckRanking(Battle, Atacante, UserList(Atacante).flags.BattlePuntos)
118         Call CheckRanking(Battle, Muerto, UserList(Muerto).flags.BattlePuntos)
120         Call GuardarRanking

        End If

        
        Exit Sub

ContarPuntoBattle_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ContarPuntoBattle", Erl)
        Resume Next
        
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef obj As obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
        
        On Error GoTo Tilelibre_Err
        

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

100     hayobj = False
102     nPos.Map = Pos.Map
    
104     Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
106         If LoopC > 15 Then
108             Notfound = True
                Exit Do

            End If
        
110         For tY = Pos.Y - LoopC To Pos.Y + LoopC
112             For tX = Pos.X - LoopC To Pos.X + LoopC
            
114                 If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                        'We continue if: a - the item is different from 0 and the dropped item or b - the Amount dropped + Amount in map exceeds MAX_INVENTORY_OBJS
116                     hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)

118                     If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.Amount + obj.Amount > MAX_INVENTORY_OBJS)

120                     If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
122                         nPos.X = tX
124                         nPos.Y = tY
126                         tX = Pos.X + LoopC
128                         tY = Pos.Y + LoopC

                        End If

                    End If
            
130             Next tX
132         Next tY
        
134         LoopC = LoopC + 1
        
        Loop
    
136     If Notfound = True Then
138         nPos.X = 0
140         nPos.Y = 0

        End If

        
        Exit Sub

Tilelibre_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.Tilelibre", Erl)
        Resume Next
        
End Sub

Sub WarpToLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal FX As Boolean = False)
        'Santo: Sub para buscar la posición legal mas cercana al objetivo y warpearlo.
        
        On Error GoTo WarpToLegalPos_Err
        

        Dim ALoop As Byte, Find As Boolean, lX As Long, lY As Long

100     Find = False
102     ALoop = 1

104     Do Until Find = True

106         For lX = X - ALoop To X + ALoop
108             For lY = Y - ALoop To Y + ALoop

110                 With MapData(Map, lX, lY)

112                     If .UserIndex <= 0 Then
114                         If (.Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Then
116                             If .TileExit.Map = 0 Then
118                                 If .NpcIndex <= 0 Then
120                                     Call WarpUserChar(UserIndex, Map, lX, lY, FX)
122                                     Find = True
                                        Exit Sub

                                    End If

                                End If

                            End If

                        End If

                    End With

124             Next lY
126         Next lX

128         ALoop = ALoop + 1
        Loop

        
        Exit Sub

WarpToLegalPos_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.WarpToLegalPos", Erl)
        Resume Next
        
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
        
        On Error GoTo WarpUserChar_Err
        

        Dim OldMap As Integer

        Dim OldX   As Integer

        Dim OldY   As Integer
    
100     If UserList(UserIndex).ComUsu.DestUsu > 0 Then
102         If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
104             If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
106                 Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
108                 Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                

                End If

            End If

        End If
    
        'Quitar el dialogo
110     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    
112     Call WriteRemoveAllDialogs(UserIndex)
    
114     OldMap = UserList(UserIndex).Pos.Map
116     OldX = UserList(UserIndex).Pos.X
118     OldY = UserList(UserIndex).Pos.Y
    
120     Call EraseUserChar(UserIndex, True)
    
122     If OldMap <> Map Then
        
124         Call WriteChangeMap(UserIndex, Map)
            'Call WriteLight(UserIndex, map)
126         'Call WriteHora(UserIndex)
        
            ' If MapInfo(OldMap).music_numberLow <> MapInfo(map).music_numberLow Then
            'Call WritePlayMidi(UserIndex, MapInfo(map).music_numberLow, 1)
            'End If
        
128         If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And UserList(UserIndex).Stats.ELV < 42 Then
130             Call WriteConsoleMsg(UserIndex, "Estas saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.", FontTypeNames.FONTTYPE_WARNING)

            End If
        
132         UserList(UserIndex).flags.NecesitaOxigeno = RequiereOxigeno(Map)

134         If UserList(UserIndex).flags.NecesitaOxigeno Then
136             Call WriteContadores(UserIndex)
138             Call WriteOxigeno(UserIndex)

140             If UserList(UserIndex).Counters.Oxigeno = 0 Then
142                 UserList(UserIndex).flags.Ahogandose = 1

                End If

            End If
            
            UserList(UserIndex).Counters.TiempoDeInmunidad = INTERVALO_INMUNIDAD
            UserList(UserIndex).flags.Inmunidad = 1

144         If RequiereOxigeno(OldMap) = True And UserList(UserIndex).flags.NecesitaOxigeno = False Then  'And UserList(UserIndex).Stats.ELV < 35 Then
        
                'Call WriteConsoleMsg(UserIndex, "Ya no necesitas oxigeno.", FontTypeNames.FONTTYPE_WARNING)
146             Call WriteContadores(UserIndex)
148             Call WriteOxigeno(UserIndex)

            End If
        
            'Update new Map Users
150         MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
            'Update old Map Users
152         MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

154         If MapInfo(OldMap).NumUsers < 0 Then
156             MapInfo(OldMap).NumUsers = 0

            End If
            
            'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
            Dim nextMap, previousMap As Boolean
            
            nextMap = distanceToCities(Map).distanceToCity(UserList(UserIndex).Hogar) >= 0
            previousMap = distanceToCities(UserList(UserIndex).Pos.Map).distanceToCity(UserList(UserIndex).Hogar) >= 0

            If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            
            ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
                UserList(UserIndex).flags.lastMap = UserList(UserIndex).Pos.Map
            
            ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el ultimo mapa es 0 ya que no esta en un dungeon)
                UserList(UserIndex).flags.lastMap = 0
            
            ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
                UserList(UserIndex).flags.lastMap = UserList(UserIndex).flags.lastMap

            End If
        

            If UserList(UserIndex).flags.Traveling = 1 Then
                UserList(UserIndex).flags.Traveling = 0
                UserList(UserIndex).Counters.goHome = 0
                Call WriteConsoleMsg(UserIndex, "El viaje ha terminado", FontTypeNames.FONTTYPE_INFOBOLD)
    
            End If

        End If
    
158     UserList(UserIndex).Pos.X = X
160     UserList(UserIndex).Pos.Y = Y
162     UserList(UserIndex).Pos.Map = Map
    
164     If FX Then
166         Call MakeUserChar(True, Map, UserIndex, Map, X, Y, 1)
        Else
168         Call MakeUserChar(True, Map, UserIndex, Map, X, Y, 0)

        End If
    
170     Call WriteUserCharIndexInServer(UserIndex)
    
        'Force a flush, so user index is in there before it's destroyed for teleporting
    
    
        'Seguis invisible al pasar de mapa
172     If (UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
174         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))

        End If
    
        'Reparacion temporal del bug de particulas. 08/07/09 LADDER

        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 71, False))
    
        If UserList(UserIndex).flags.AdminInvisible = 0 Then
176         If FX Then 'FX
178             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
180             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))
            End If
        Else
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
        End If
        
        If UserList(UserIndex).NroMascotas > 0 Then Call WarpMascotas(UserIndex)
    
182     If MapInfo(Map).zone = "DUNGEON" Then
184         If UserList(UserIndex).flags.Montado > 0 Then
186             Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

            End If

        End If
    
        ' Call WarpFamiliar(UserIndex)
        
        Exit Sub

WarpUserChar_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.WarpUserChar", Erl)
        Resume Next
        
End Sub

Sub WarpFamiliar(ByVal UserIndex As Integer)
        
        On Error GoTo WarpFamiliar_Err
        

100     With UserList(UserIndex)

102         If .Familiar.Invocado = 1 Then
104             Call QuitarNPC(.Familiar.Id)
                ' If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
106             .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, UserList(UserIndex).Pos, False, True)

                'Controlamos que se sumoneo OK
108             If .Familiar.Id = 0 Then
110                 Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota. Se provoco un ERROR.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

112             Call CargarFamiliar(UserIndex)
            Else

                'Call WriteConsoleMsg(UserIndex, "No se permiten familiares en zona segura. " & .Familiar.Nombre & " te esperará afuera.", FontTypeNames.FONTTYPE_INFO)
            End If
    
        End With
            
        
        Exit Sub

WarpFamiliar_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.WarpFamiliar", Erl)
        Resume Next
        
End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer)
        
        On Error GoTo Cerrar_Usuario_Err
        
 
100     If UserList(UserIndex).flags.UserLogged And Not UserList(UserIndex).Counters.Saliendo Then
102         UserList(UserIndex).Counters.Saliendo = True

104         If UserList(UserIndex).flags.Privilegios = PlayerType.user And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 And UserList(UserIndex).flags.Muerto = 0 Then
106             UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
108             Call WriteLocaleMsg(UserIndex, "203", FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Counters.Salir)
                'Call WriteConsoleMsg(UserIndex, "Saliendo...Se saldrá del juego en " & UserList(UserIndex).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
            Else
            
                'Call WriteConsoleMsg(UserIndex, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteDisconnect(UserIndex)
            
  
112             Call CloseSocket(UserIndex)

            End If

        End If

        
        Exit Sub

Cerrar_Usuario_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.Cerrar_Usuario", Erl)
        Resume Next
        
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal UserIndex As Integer)
        
        On Error GoTo CancelExit_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/02/08
        '
        '***************************************************
100     If UserList(UserIndex).Counters.Saliendo Then

            ' Is the user still connected?
102         If UserList(UserIndex).ConnIDValida Then
104             UserList(UserIndex).Counters.Saliendo = False
106             UserList(UserIndex).Counters.Salir = 0
108             Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Else

                'Simply reset
110             If UserList(UserIndex).flags.Privilegios = PlayerType.user And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 Then
112                 UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
                Else
114                 Call WriteConsoleMsg(UserIndex, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
116                 Call WriteDisconnect(UserIndex)
                
118                 Call CloseSocket(UserIndex)

                End If
            
                'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
            End If

        End If

        
        Exit Sub

CancelExit_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.CancelExit", Erl)
        Resume Next
        
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal UserIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
        
        On Error GoTo CambiarNick_Err
        

        Dim ViejoNick       As String

        Dim ViejoCharBackup As String

100     If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
102     ViejoNick = UserList(UserIndexDestino).name

104     If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
            'hace un backup del char
106         ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
108         Name CharPath & ViejoNick & ".chr" As ViejoCharBackup

        End If

        
        Exit Sub

CambiarNick_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.CambiarNick", Erl)
        Resume Next
        
End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal nombre As String)
        
        On Error GoTo SendUserStatsTxtOFF_Err
        

100     If FileExist(CharPath & nombre & ".chr", vbArchive) = False Then
102         Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Else
104         Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & nombre, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Vitalidad: " & GetVar(CharPath & nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
    
112         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
    
114         Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
116         Call WriteConsoleMsg(sendIndex, "Veces Que Murio: " & GetVar(CharPath & nombre & ".chr", "Flags", "VecesQueMoriste"), FontTypeNames.FONTTYPE_INFO)
            #If ConUpTime Then

                Dim TempSecs As Long

                Dim TempStr  As String

118             TempSecs = GetVar(CharPath & nombre & ".chr", "INIT", "UpTime")
120             TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
122             Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
            #End If

        End If

        
        Exit Sub

SendUserStatsTxtOFF_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserStatsTxtOFF", Erl)
        Resume Next
        
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
        
        On Error GoTo VolverCriminal_Err
        

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/06/2006
        'Nacho: Actualiza el tag al cliente
        '**************************************************************
100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

102     If UserList(UserIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
   
104         If UserList(UserIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)

        End If

106     If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Exit Sub

108     UserList(UserIndex).Faccion.Status = 0

110     Call RefreshCharStatus(UserIndex)

        
        Exit Sub

VolverCriminal_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.VolverCriminal", Erl)
        Resume Next
        
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/06/2006
        'Nacho: Actualiza el tag al cliente.
        '**************************************************************
        
        On Error GoTo VolverCiudadano_Err
        

100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

102     UserList(UserIndex).Faccion.Status = 1
104     Call RefreshCharStatus(UserIndex)

        
        Exit Sub

VolverCiudadano_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.VolverCiudadano", Erl)
        Resume Next
        
End Sub

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
        '***************************************************
        'Author: Unknown
        'Last Modification: 30/09/2020
        '
        '***************************************************
        
        On Error GoTo getMaxInventorySlots_Err
        

100     If UserList(UserIndex).Stats.InventLevel > 0 Then
102         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS + UserList(UserIndex).Stats.InventLevel * SLOTS_PER_ROW_INVENTORY
        Else
104         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS

        End If

        
        Exit Function

getMaxInventorySlots_Err:
        Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.getMaxInventorySlots", Erl)
        Resume Next
        
End Function

Private Sub WarpMascotas(ByVal UserIndex As Integer)

    '************************************************
    'Author: Uknown
    'Last Modified: 26/10/2010
    '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
    '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
    '11/05/2009: ZaMa - Chequeo si la mascota pueden spawnear para asiganrle los stats.
    '26/10/2010: ZaMa - Ahora las mascotas respawnean de forma aleatoria.
    '************************************************
    Dim i                As Integer

    Dim petType          As Integer

    Dim PetRespawn       As Boolean

    Dim PetTiempoDeVida  As Integer

    Dim canWarp          As Boolean

    Dim Index            As Integer

    Dim iMinHP           As Integer

    canWarp = (Not MapInfo(UserList(UserIndex).Pos.Map).Seguro)

    For i = 1 To MAXMASCOTAS
        Index = UserList(UserIndex).MascotasIndex(i)
        
        If Index > 0 Then
            'Store data and remove NPC to recreate it after warp
            petType = UserList(UserIndex).MascotasType(i)
            PetTiempoDeVida = Npclist(Index).Contadores.TiempoExistencia
            
            ' Guardamos el hp, para restaurarlo cuando se cree el npc
            iMinHP = Npclist(Index).Stats.MinHp
            
            Call QuitarNPC(Index)
            
            ' Restauramos el valor de la variable
            UserList(UserIndex).MascotasType(i) = petType
            
            If petType > 0 And canWarp And UserList(UserIndex).flags.MascotasGuardadas = 0 Then
        
                Dim SpawnPos As WorldPos
            
                SpawnPos.Map = UserList(UserIndex).Pos.Map
                SpawnPos.X = UserList(UserIndex).Pos.X + RandomNumber(-3, 3)
                SpawnPos.Y = UserList(UserIndex).Pos.Y + RandomNumber(-3, 3)
            
                Index = SpawnNpc(petType, SpawnPos, False, PetRespawn)
                
                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
                If Index > 0 Then
                    UserList(UserIndex).MascotasIndex(i) = Index
    
                    ' Nos aseguramos de que conserve el hp, si estaba danado
                    Npclist(Index).Stats.MinHp = IIf(iMinHP = 0, Npclist(Index).Stats.MinHp, iMinHP)
                
                    Npclist(Index).MaestroUser = UserIndex
                    Npclist(Index).Contadores.TiempoExistencia = PetTiempoDeVida
                    Call FollowAmo(Index)
    
                End If
    
            End If
            
        End If

    Next i
    
    If Not canWarp And UserList(UserIndex).flags.MascotasGuardadas = 0 Then
        Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Estas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub
