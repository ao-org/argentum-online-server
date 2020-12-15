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
138     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ActStats", Erl)
140     Resume Next
        
End Sub

Sub RevivirUsuario(ByVal Userindex As Integer)
        
        On Error GoTo RevivirUsuario_Err
        
        With UserList(Userindex)

104         .flags.Muerto = 0
106         .Stats.MinHp = .Stats.MaxHp

            Call WriteUpdateHP(Userindex)
    
112         If .flags.Navegando = 1 Then
    
                Dim Barco As ObjData
    
114             Barco = ObjData(.Invent.BarcoObjIndex)

122             If Barco.Subtipo = 0 Then
124                 .Char.Head = .OrigChar.Head
                    
                    If .Invent.CascoEqpObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                    End If

126                 Call WriteNadarToggle(Userindex, True)
                Else
                    .Char.Head = 0
128                 Call WriteNadarToggle(Userindex, False)
                End If
        
                .Char.Body = Barco.Ropaje
        
158             .Char.ShieldAnim = NingunEscudo
160             .Char.WeaponAnim = NingunArma
162             .Char.speeding = Barco.Velocidad
       
            Else
        
166             .Char.Head = .OrigChar.Head
    
168             If .Invent.CascoEqpObjIndex > 0 Then
170                 .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                End If
    
172             If .Invent.EscudoEqpObjIndex > 0 Then
174                 .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
    
                End If
    
176             If .Invent.WeaponEqpObjIndex > 0 Then
178                 .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
        
180                 If ObjData(.Invent.WeaponEqpObjIndex).CreaGRH <> "" Then
182                     .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH
184                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
    
                    End If
            
                End If
    
186             If .Invent.ArmourEqpObjIndex > 0 Then
188                 .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        
190                 If ObjData(.Invent.ArmourEqpObjIndex).CreaGRH <> "" Then
192                     .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH
194                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))
    
                    End If
    
                Else
196                 Call DarCuerpoDesnudo(Userindex)
            
                End If
    
198             If .Invent.EscudoEqpObjIndex > 0 Then
200                 .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
    
202                 If ObjData(.Invent.EscudoEqpObjIndex).CreaGRH <> "" Then
204                     .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH
206                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
    
                    End If
            
                End If
    
208             If .Invent.CascoEqpObjIndex > 0 Then
210                 .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
    
212                 If ObjData(.Invent.CascoEqpObjIndex).CreaGRH <> "" Then
214                     .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH
216                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
    
                    End If
            
                End If
    
218             If .Invent.MagicoObjIndex > 0 Then
220                 If ObjData(.Invent.MagicoObjIndex).CreaGRH <> "" Then
222                     .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
224                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
    
                    End If
    
                End If
    
226             If .Invent.NudilloObjIndex > 0 Then
228                 If ObjData(.Invent.NudilloObjIndex).CreaGRH <> "" Then
230                     .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
232                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
    
                    End If
                End If
                
234             If .Invent.AnilloEqpObjIndex > 0 Then
236                 If ObjData(.Invent.AnilloEqpObjIndex).CreaGRH <> "" Then
238                     .Char.Anillo_Aura = ObjData(.Invent.AnilloEqpObjIndex).CreaGRH
240                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Anillo_Aura, False, 6))
                    End If
                End If
                
                .Char.speeding = VelocidadNormal
    
            End If
    
242         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))

        End With
        
        Exit Sub

RevivirUsuario_Err:
246     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.RevivirUsuario", Erl)
248     Resume Next
        
End Sub

Sub ChangeUserChar(ByVal Userindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
        
        On Error GoTo ChangeUserChar_Err
        

100     With UserList(Userindex).Char
102         .Body = Body
104         .Head = Head
106         .Heading = Heading
108         .WeaponAnim = Arma
110         .ShieldAnim = Escudo
112         .CascoAnim = Casco

        End With
    
114     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterChange(Body, Head, Heading, UserList(Userindex).Char.CharIndex, Arma, Escudo, UserList(Userindex).Char.FX, UserList(Userindex).Char.loops, Casco, False))

        
        Exit Sub

ChangeUserChar_Err:
116     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ChangeUserChar", Erl)
118     Resume Next
        
End Sub

Sub EraseUserChar(ByVal Userindex As Integer, ByVal Desvanecer As Boolean)

        On Error GoTo ErrorHandler

        Dim Error As String
   
100     Error = "1"

102     If UserList(Userindex).Char.CharIndex = 0 Then Exit Sub
   
104     CharList(UserList(Userindex).Char.CharIndex) = 0
    
106     If UserList(Userindex).Char.CharIndex = LastChar Then

108         Do Until CharList(LastChar) > 0
110             LastChar = LastChar - 1

112             If LastChar <= 1 Then Exit Do
            Loop

        End If

114     Error = "2"
    
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
116     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterRemove(UserList(Userindex).Char.CharIndex, Desvanecer))
118     Error = "3"
120     Call QuitarUser(Userindex, UserList(Userindex).Pos.Map)
122     Error = "4"
    
124     MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).Userindex = 0
126     Error = "5"
128     UserList(Userindex).Char.CharIndex = 0
    
130     NumChars = NumChars - 1
132     Error = "6"
        Exit Sub
    
ErrorHandler:
134     Call LogError("Error en EraseUserchar " & Error & " - " & Err.Number & ": " & Err.description)

End Sub

Sub RefreshCharStatus(ByVal Userindex As Integer)
        
        On Error GoTo RefreshCharStatus_Err
        

        '*************************************************
        'Author: Tararira
        'Last modified: 6/04/2007
        'Refreshes the status and tag of UserIndex.
        '*************************************************
        Dim klan As String

100     If UserList(Userindex).GuildIndex > 0 Then
102         klan = modGuilds.GuildName(UserList(Userindex).GuildIndex)
104         klan = " <" & klan & ">"

        End If
    
106     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTagAndStatus(Userindex, UserList(Userindex).Faccion.Status, UserList(Userindex).name & klan))

        
        Exit Sub

RefreshCharStatus_Err:
108     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.RefreshCharStatus", Erl)
110     Resume Next
        
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal appear As Byte = 0)

        On Error GoTo hayerror

        Dim CharIndex As Integer

        Dim errort    As String

100     If InMapBounds(Map, X, Y) Then

            'If needed make a new character in list
102         If UserList(Userindex).Char.CharIndex = 0 Then
104             CharIndex = NextOpenCharIndex
106             UserList(Userindex).Char.CharIndex = CharIndex
108             CharList(CharIndex) = Userindex
            End If

110         errort = "1"
        
            'Place character on map if needed
112         If toMap Then MapData(Map, X, Y).Userindex = Userindex
114         errort = "2"
        
            'Send make character command to clients
            Dim klan       As String

            Dim clan_nivel As Byte

116         If UserList(Userindex).GuildIndex > 0 Then
118             klan = modGuilds.GuildName(UserList(Userindex).GuildIndex)
120             clan_nivel = modGuilds.NivelDeClan(UserList(Userindex).GuildIndex)
            End If

122         errort = "3"
        
            Dim bCr As Byte
        
124         bCr = UserList(Userindex).Faccion.Status
126         errort = "4"
        
128         If LenB(klan) <> 0 Then
130             If Not toMap Then
132                 errort = "5"
134                 Call WriteCharacterCreate(sndIndex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.CharIndex, X, Y, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.FX, 999, UserList(Userindex).Char.CascoAnim, UserList(Userindex).name & " <" & klan & ">", bCr, UserList(Userindex).flags.Privilegios, UserList(Userindex).Char.ParticulaFx, UserList(Userindex).Char.Head_Aura, UserList(Userindex).Char.Arma_Aura, UserList(Userindex).Char.Body_Aura, UserList(Userindex).Char.Anillo_Aura, UserList(Userindex).Char.Otra_Aura, UserList(Userindex).Char.Escudo_Aura, UserList(Userindex).Char.speeding, False, UserList(Userindex).donador.activo, appear, UserList(Userindex).Grupo.Lider, UserList(Userindex).GuildIndex, clan_nivel, UserList(Userindex).Stats.MinHp, UserList(Userindex).Stats.MaxHp, 0)
                Else
136                 errort = "6"
138                 Call AgregarUser(Userindex, UserList(Userindex).Pos.Map, appear)
                End If

            Else 'if tiene clan

140             If Not toMap Then
142                 errort = "7"
144                 Call WriteCharacterCreate(sndIndex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.CharIndex, X, Y, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.FX, 999, UserList(Userindex).Char.CascoAnim, UserList(Userindex).name, bCr, UserList(Userindex).flags.Privilegios, UserList(Userindex).Char.ParticulaFx, UserList(Userindex).Char.Head_Aura, UserList(Userindex).Char.Arma_Aura, UserList(Userindex).Char.Body_Aura, UserList(Userindex).Char.Anillo_Aura, UserList(Userindex).Char.Otra_Aura, UserList(Userindex).Char.Escudo_Aura, UserList(Userindex).Char.speeding, False, UserList(Userindex).donador.activo, appear, UserList(Userindex).Grupo.Lider, 0, 0, UserList(Userindex).Stats.MinHp, UserList(Userindex).Stats.MaxHp, 0)
                Else
146                 errort = "8"
148                 Call AgregarUser(Userindex, UserList(Userindex).Pos.Map, appear)
                End If

            End If 'if clan

        End If

        Exit Sub

hayerror:
150     LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description & " - Nombre del usuario " & UserList(Userindex).name) & " - " & errort & "- Pos: M: " & Map & " X: " & X & " Y: " & Y
        'Resume Next
152     Call CloseSocket(Userindex)

End Sub

Sub CheckUserLevel(ByVal Userindex As Integer)
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
    
100     With UserList(Userindex)

102         WasNewbie = EsNewbie(Userindex)
        
104         Do While .Stats.Exp >= .Stats.ELU And .Stats.ELV < STAT_MAXELV
            
                'Store it!
                'Call Statistics.UserLevelUp(UserIndex)

106             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 106, 0))
108             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
110             Call WriteLocaleMsg(Userindex, "186", FontTypeNames.FONTTYPE_INFO)
            
112             Pts = Pts + 5
            
114             .Stats.ELV = .Stats.ELV + 1
            
116             .Stats.Exp = .Stats.Exp - .Stats.ELU

118             If .Stats.ELV < 15 Then
120                 .Stats.ELU = .Stats.ELU * 1.4
122             ElseIf .Stats.ELV < 21 Then
124                 .Stats.ELU = .Stats.ELU * 1.35
126             ElseIf .Stats.ELV < 26 Then
128                 .Stats.ELU = .Stats.ELU * 1.3
130             ElseIf .Stats.ELV < 35 Then
132                 .Stats.ELU = .Stats.ELU * 1.2
134             ElseIf .Stats.ELV < 40 Then
136                 .Stats.ELU = .Stats.ELU * 1.3
                Else
138                 .Stats.ELU = .Stats.ELU * 1.375

                End If
            
                'Calculo subida de vida
140             Promedio = ModVida(.clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
142             aux = RandomNumber(0, 100)
            
144             If Promedio - Int(Promedio) = 0.5 Then
                    'Es promedio semientero
146                 DistVida(1) = DistribucionSemienteraVida(1)
148                 DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
150                 DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
152                 DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
154                 If aux <= DistVida(1) Then
156                     AumentoHP = Promedio + 1.5
158                 ElseIf aux <= DistVida(2) Then
160                     AumentoHP = Promedio + 0.5
162                 ElseIf aux <= DistVida(3) Then
164                     AumentoHP = Promedio - 0.5
                    Else
166                     AumentoHP = Promedio - 1.5

                    End If

                Else
                    'Es promedio entero
168                 DistVida(1) = DistribucionEnteraVida(1)
170                 DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
172                 DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
174                 DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
176                 DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
178                 If aux <= DistVida(1) Then
180                     AumentoHP = Promedio + 2
182                 ElseIf aux <= DistVida(2) Then
184                     AumentoHP = Promedio + 1
186                 ElseIf aux <= DistVida(3) Then
188                     AumentoHP = Promedio
190                 ElseIf aux <= DistVida(4) Then
192                     AumentoHP = Promedio - 1
                    Else
194                     AumentoHP = Promedio - 2

                    End If
                
                End If
            
196             Select Case .clase

                    Case eClass.Mage '
198                     AumentoHIT = 1
200                     AumentoMANA = 2.8 * .Stats.UserAtributos(eAtributos.Inteligencia)
202                     AumentoSTA = AumentoSTMago

204                 Case eClass.Bard 'Balanceda Mana
206                     AumentoHIT = 2
208                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
210                     AumentoSTA = AumentoSTDef

212                 Case eClass.Druid 'Balanceda Mana
214                     AumentoHIT = 2
216                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
218                     AumentoSTA = AumentoSTDef

220                 Case eClass.Assasin
222                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
224                     AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
226                     AumentoSTA = AumentoSTDef

228                 Case eClass.Cleric 'Balanceda Mana
230                     AumentoHIT = 2
232                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
234                     AumentoSTA = AumentoSTDef

236                 Case eClass.Paladin
238                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
240                     AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
242                     AumentoSTA = AumentoSTDef
                
244                 Case eClass.Thief
246                     AumentoHIT = 2
248                     AumentoSTA = AumentoSTLadron

250                 Case eClass.Hunter
252                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
254                     AumentoSTA = AumentoSTDef

256                 Case eClass.Trabajador
258                     AumentoHIT = 2
260                     AumentoSTA = AumentoSTDef + 5
                    
262                 Case eClass.Pirat
264                     AumentoHIT = 3
266                     AumentoSTA = AumentoSTDef
                    
268                 Case eClass.Bandit
270                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
272                     AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
274                     AumentoSTA = AumentoStBandido

276                 Case eClass.Warrior
278                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
280                     AumentoSTA = AumentoSTDef

282                 Case Else
284                     AumentoHIT = 2
286                     AumentoSTA = AumentoSTDef

                End Select
            
                'Actualizamos HitPoints
288             .Stats.MaxHp = .Stats.MaxHp + AumentoHP

290             If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                'Actualizamos Stamina
292             .Stats.MaxSta = .Stats.MaxSta + AumentoSTA

294             If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
                'Actualizamos Mana
296             .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

298             If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

                'Actualizamos Golpe Máximo
300             .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            
                'Actualizamos Golpe Mínimo
302             .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
        
                'Notificamos al user
304             If AumentoHP > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
306                 Call WriteLocaleMsg(Userindex, "197", FontTypeNames.FONTTYPE_INFO, AumentoHP)

                End If

308             If AumentoSTA > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
310                 Call WriteLocaleMsg(Userindex, "198", FontTypeNames.FONTTYPE_INFO, AumentoSTA)

                End If

312             If AumentoMANA > 0 Then
314                 Call WriteLocaleMsg(Userindex, "199", FontTypeNames.FONTTYPE_INFO, AumentoMANA)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
                End If

316             If AumentoHIT > 0 Then
318                 Call WriteLocaleMsg(Userindex, "200", FontTypeNames.FONTTYPE_INFO, AumentoHIT)

                    'Call WriteConsoleMsg(UserIndex, "Tu golpe aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                End If

320             PasoDeNivel = True
             
                ' Call LogDesarrollo(.name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)

322             .Stats.MinHp = .Stats.MaxHp
            
                ' Call UpdateUserInv(True, UserIndex, 0)
            
324             If OroPorNivel > 0 Then
326                 If EsNewbie(Userindex) Then
                        Dim OroRecompenza As Long
    
328                     OroRecompenza = OroPorNivel * .Stats.ELV * OroMult * .flags.ScrollOro
330                     .Stats.GLD = .Stats.GLD + OroRecompenza
                        'Call WriteConsoleMsg(UserIndex, "Has ganado " & OroRecompenza & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
332                     Call WriteLocaleMsg(Userindex, "29", FontTypeNames.FONTTYPE_INFO, OroRecompenza)
                    End If
                End If
            
334             If Not EsNewbie(Userindex) And WasNewbie Then
        
336                 Call QuitarNewbieObj(Userindex)
            
                End If
        
            Loop
        
338         If PasoDeNivel Then
                'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo nivel
340             If .Stats.ELV >= STAT_MAXELV Then
342                 .Stats.Exp = 0
344                 .Stats.ELU = 0
                End If
        
346             Call UpdateUserInv(True, Userindex, 0)
                'Call CheckearRecompesas(UserIndex, 3)
348             Call WriteUpdateUserStats(Userindex)
            
350             If Pts > 0 Then
                
352                 .Stats.SkillPts = .Stats.SkillPts + Pts
354                 Call WriteLevelUp(Userindex, .Stats.SkillPts)
356                 Call WriteLocaleMsg(Userindex, "187", FontTypeNames.FONTTYPE_INFO, Pts)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
                End If

            End If
    
        End With
    
        Exit Sub

ErrHandler:
358     Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub

Function PuedeAtravesarAgua(ByVal Userindex As Integer) As Boolean
        
        On Error GoTo PuedeAtravesarAgua_Err
        

100     PuedeAtravesarAgua = UserList(Userindex).flags.Navegando = 1
        'If PuedeAtravesarAgua = True Then
        '  Exit Function
        'Else
        '  If UserList(UserIndex).flags.Nadando = 1 Then
        ' PuedeAtravesarAgua = True
        'End If
        'End If

        
        Exit Function

PuedeAtravesarAgua_Err:
102     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.PuedeAtravesarAgua", Erl)
104     Resume Next
        
End Function

Sub MoveUserChar(ByVal Userindex As Integer, ByVal nHeading As eHeading)
        
        On Error GoTo MoveUserChar_Err
        

        Dim nPos         As WorldPos

        Dim nPosOriginal As WorldPos

        Dim nPosMuerto   As WorldPos

        Dim sailing      As Boolean

100     With UserList(Userindex)

102         If .Accion.AccionPendiente = True Then
104             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(.Char.CharIndex, .Accion.Particula, 1, True))
106             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(.Char.CharIndex, 1, Accion_Barra.CancelarAccion))
108             .Accion.AccionPendiente = False
110             .Accion.Particula = 0
112             .Accion.TipoAccion = Accion_Barra.CancelarAccion
114             .Accion.HechizoPendiente = 0
116             .Accion.RunaObj = 0
118             .Accion.ObjSlot = 0
120             .Accion.AccionPendiente = False

            End If

122         sailing = PuedeAtravesarAgua(Userindex)
124         nPos = .Pos
126         Call HeadtoPos(nHeading, nPos)
        
128         If MapData(nPos.Map, nPos.X, nPos.Y).TileExit.Map <> 0 And .Counters.TiempoDeMapeo > 0 Then
130             If .flags.Muerto = 0 Then
132                 Call WriteConsoleMsg(Userindex, "Estas en combate, debes aguardar " & .Counters.TiempoDeMapeo & " segundo(s) para escapar...", FontTypeNames.FONTTYPE_INFOBOLD)
134                 Call WritePosUpdate(Userindex)
                    Exit Sub
    
                End If
    
            End If
    
136         If MapData(nPos.Map, nPos.X, nPos.Y).Userindex <> 0 Then
                Dim IndexMuerto As Integer
138             IndexMuerto = MapData(nPos.Map, nPos.X, nPos.Y).Userindex

140             If UserList(IndexMuerto).flags.Muerto = 1 Or UserList(IndexMuerto).flags.AdminInvisible = 1 Then

142                 Call WarpToLegalPos(IndexMuerto, UserList(IndexMuerto).Pos.Map, UserList(IndexMuerto).Pos.X, UserList(IndexMuerto).Pos.Y, False)
    
                Else
144                 Call WritePosUpdate(Userindex)
    
                    'Call WritePosUpdate(MapData(nPos.Map, nPos.X, nPos.Y).UserIndex)
                End If
    
            End If
    
146         If LegalWalk(.Pos.Map, nPos.X, nPos.Y, nHeading, sailing, Not sailing, .flags.Montado) Then
148             If MapInfo(.Pos.Map).NumUsers > 1 Then
                    'si no estoy solo en el mapa...
    
150                 If .flags.AdminInvisible = 0 Then
152                     Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                    Else
154                     Call SendData(SendTarget.ToAdminAreaButIndex, Userindex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                    End If
            
                End If
    
                'Call RefreshAllUser(UserIndex) '¿Clones? Ladder probar
                'Update map and user pos
156             MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex = 0
158             .Pos = nPos
160             .Char.Heading = nHeading
162             MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex = Userindex
            
                'Actualizamos las áreas de ser necesario
164             Call ModAreas.CheckUpdateNeededUser(Userindex, nHeading, 0)
           
            Else
166             Call WritePosUpdate(Userindex)
    
            End If
        
168         If .Counters.Trabajando Then
170             Call WriteMacroTrabajoToggle(Userindex, False)
    
            End If
    
172         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

        End With

        
        Exit Sub

MoveUserChar_Err:
174     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.MoveUserChar", Erl)
176     Resume Next
        
End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

        '*************************************************
        'Author: ZaMa
        'Last modified: 30/03/2009
        'Returns the heading opposite to the one passed by val.
        '*************************************************
100     Select Case nHeading

            Case eHeading.EAST
102             InvertHeading = eHeading.WEST

104         Case eHeading.WEST
106             InvertHeading = eHeading.EAST

108         Case eHeading.SOUTH
110             InvertHeading = eHeading.NORTH

112         Case eHeading.NORTH
114             InvertHeading = eHeading.SOUTH

        End Select

End Function

Sub ChangeUserInv(ByVal Userindex As Integer, ByVal slot As Byte, ByRef Object As UserOBJ)
        
        On Error GoTo ChangeUserInv_Err
        
100     UserList(Userindex).Invent.Object(slot) = Object
102     Call WriteChangeInventorySlot(Userindex, slot)

        
        Exit Sub

ChangeUserInv_Err:
104     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ChangeUserInv", Erl)
106     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.NextOpenCharIndex", Erl)
114     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.NextOpenUser", Erl)
112     Resume Next
        
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
        
        On Error GoTo SendUserStatsTxt_Err
        

        Dim GuildI As Integer

100     Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(Userindex).Stats.ELV & "  EXP: " & UserList(Userindex).Stats.Exp & "/" & UserList(Userindex).Stats.ELU, FontTypeNames.FONTTYPE_INFO)
104     Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(Userindex).Stats.MinHp & "/" & UserList(Userindex).Stats.MaxHp & "  Mana: " & UserList(Userindex).Stats.MinMAN & "/" & UserList(Userindex).Stats.MaxMAN & "  Vitalidad: " & UserList(Userindex).Stats.MinSta & "/" & UserList(Userindex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
106     If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
108         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(Userindex).Stats.MinHIT & "/" & UserList(Userindex).Stats.MaxHit & " (" & ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(Userindex).Stats.MinHIT & "/" & UserList(Userindex).Stats.MaxHit, FontTypeNames.FONTTYPE_INFO)

        End If
    
112     If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
114         If UserList(Userindex).Invent.EscudoEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(Userindex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
118             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
120         Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
    
122     If UserList(Userindex).Invent.CascoEqpObjIndex > 0 Then
124         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(Userindex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
126         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
    
128     GuildI = UserList(Userindex).GuildIndex

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

138         TempDate = Now - UserList(Userindex).LogOnTime
140         TempSecs = (UserList(Userindex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
142         TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
144         Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
146         Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If

148     Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(Userindex).Stats.GLD & "  Posicion: " & UserList(Userindex).Pos.X & "," & UserList(Userindex).Pos.Y & " en mapa " & UserList(Userindex).Pos.Map, FontTypeNames.FONTTYPE_INFO)
150     Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion) & ", " & UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma), FontTypeNames.FONTTYPE_INFO)
152     Call WriteConsoleMsg(sendIndex, "Veces que Moriste: " & UserList(Userindex).flags.VecesQueMoriste, FontTypeNames.FONTTYPE_INFO)

        Exit Sub

SendUserStatsTxt_Err:
154     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserStatsTxt", Erl)
156     Resume Next
        
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
        
        On Error GoTo SendUserMiniStatsTxt_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is online.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
100     With UserList(Userindex)
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
126     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)
128     Resume Next
        
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
136     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserMiniStatsTxtFromChar", Erl)
138     Resume Next
        
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim j As Long
    
100     Call WriteConsoleMsg(sendIndex, UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(Userindex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
104     For j = 1 To UserList(Userindex).CurrentInventorySlots

106         If UserList(Userindex).Invent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(Userindex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(Userindex).Invent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

            End If

110     Next j

End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

        On Error Resume Next

        Dim j        As Long

        Dim CharFile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long
    
100     CharFile = CharPath & CharName & ".chr"
    
102     If FileExist(CharFile, vbNormal) Then
104         Call WriteConsoleMsg(sendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
108         For j = 1 To MAX_INVENTORY_SLOTS
110             Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
112             ObjInd = ReadField(1, Tmp, Asc("-"))
114             ObjCant = ReadField(2, Tmp, Asc("-"))

116             If ObjInd > 0 Then
118                 Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

                End If

120         Next j

        Else
122         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)

        End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)

102     For j = 1 To NUMSKILLS
104         Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(Userindex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
        Next
106     Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(Userindex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

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
112     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.DameUserIndex", Erl)
114     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.DameUserIndexConNombre", Erl)
116     Resume Next
        
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
        
        On Error GoTo NPCAtacado_Err
        

        '**********************************************
        'Author: Unknown
        'Last Modification: 24/07/2007
        '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
        '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
        '**********************************************
        
        ' WyroX: El usuario pierde la protección
100     UserList(Userindex).Counters.TiempoDeInmunidad = 0
102     UserList(Userindex).flags.Inmunidad = 0
        
        Dim EraCriminal As Byte

        'Guardamos el usuario que ataco el npc.
104     If Npclist(NpcIndex).Movement <> ESTATICO And Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
106         Npclist(NpcIndex).Target = Userindex
108         Npclist(NpcIndex).Movement = TipoAI.NpcMaloAtacaUsersBuenos
110         Npclist(NpcIndex).flags.AttackedBy = UserList(Userindex).name
        End If

        'Npc que estabas atacando.
        Dim LastNpcHit As Integer

112     LastNpcHit = UserList(Userindex).flags.NPCAtacado
        'Guarda el NPC que estas atacando ahora.
114     UserList(Userindex).flags.NPCAtacado = NpcIndex

        'Revisamos robo de npc.
        'Guarda el primer nick que lo ataca.
116     If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then

            'El que le pegabas antes ya no es tuyo
118         If LastNpcHit <> 0 Then
120             If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(Userindex).name Then
122                 Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

                End If

            End If

124         Npclist(NpcIndex).flags.AttackedFirstBy = UserList(Userindex).name
126     ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(Userindex).name Then

            'Estas robando NPC
            'El que le pegabas antes ya no es tuyo
128         If LastNpcHit <> 0 Then
130             If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(Userindex).name Then
132                 Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

                End If

            End If

        End If

        '  EraCriminal = Status(UserIndex)

134     If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
136         If Status(Userindex) = 1 Or Status(Userindex) = 3 Then
138             Call VolverCriminal(Userindex)

            End If

        End If
        
140     If Npclist(NpcIndex).MaestroUser > 0 Then
142         If Npclist(NpcIndex).MaestroUser <> Userindex Then
144             Call AllMascotasAtacanUser(Userindex, Npclist(NpcIndex).MaestroUser)
            End If
        End If

146     Call CheckPets(NpcIndex, Userindex, False)
        
        Exit Sub

NPCAtacado_Err:
148     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.NPCAtacado", Erl)
150     Resume Next
        
End Sub

Function PuedeApuñalar(ByVal Userindex As Integer) As Boolean
        
        On Error GoTo PuedeApuñalar_Err
        

100     If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
102         PuedeApuñalar = ((UserList(Userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) And (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) Or ((UserList(Userindex).clase = eClass.Assasin) And (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))
        Else
104         PuedeApuñalar = False

        End If

        
        Exit Function

PuedeApuñalar_Err:
106     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.PuedeApuñalar", Erl)
108     Resume Next
        
End Function

Sub SubirSkill(ByVal Userindex As Integer, ByVal Skill As Integer)
        
        On Error GoTo SubirSkill_Err
        

100     If UserList(Userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

102     If UserList(Userindex).flags.Hambre = 0 And UserList(Userindex).flags.Sed = 0 Then
        
            Dim Lvl As Integer

104         Lvl = UserList(Userindex).Stats.ELV
        
106         If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
            
108         If UserList(Userindex).Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
        
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
             
132         If UserList(Userindex).flags.PendienteDelExperto = 1 Then
134             Menor = 15

            End If
        
136         If Aumenta < Menor Then
138             UserList(Userindex).Stats.UserSkills(Skill) = UserList(Userindex).Stats.UserSkills(Skill) + 1
    
140             Call WriteConsoleMsg(Userindex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(Userindex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
            
                Dim BonusExp As Long

142             BonusExp = 5 * ExpMult * UserList(Userindex).flags.ScrollExp
        
144             If UserList(Userindex).donador.activo = 1 Then
146                 BonusExp = BonusExp * 1.1

                End If
        
148             If UserList(Userindex).Stats.ELV < STAT_MAXELV Then
150                 UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + BonusExp

152                 If UserList(Userindex).Stats.Exp > MAXEXP Then UserList(Userindex).Stats.Exp = MAXEXP
            
154                 If UserList(Userindex).ChatCombate = 1 Then
156                     Call WriteLocaleMsg(Userindex, "140", FontTypeNames.FONTTYPE_EXP, BonusExp)

                    End If
                
158                 Call WriteUpdateExp(Userindex)
160                 Call CheckUserLevel(Userindex)

                End If

            End If

        End If

        
        Exit Sub

SubirSkill_Err:
162     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SubirSkill", Erl)
164     Resume Next
        
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Sub UserDie(ByVal Userindex As Integer)

        '************************************************
        'Author: Uknown
        'Last Modified: 04/15/2008 (NicoNZ)
        'Ahora se resetea el counter del invi
        '************************************************
        On Error GoTo ErrorHandler

        Dim i  As Long

        Dim aN As Integer
    
100     With UserList(Userindex)
    
            'Sonido
102         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.MUERTE_HOMBRE, .Pos.X, .Pos.Y))
        
            'Quitar el dialogo del user muerto
104         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
106         .Stats.MinHp = 0
108         .Stats.MinSta = 0
110         .flags.AtacadoPorUser = 0
112         .flags.Envenenado = 0
114         .flags.Ahogandose = 0
116         .flags.Incinerado = 0
118         .flags.incinera = 0
120         .flags.Paraliza = 0
122         .flags.Envenena = 0
124         .flags.Estupidiza = 0
126         .flags.Muerto = 1
            '.flags.SeguroParty = True
            'Call WritePartySafeOn(UserIndex)
        
128         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.MUERTE_HOMBRE, .Pos.X, .Pos.Y))
        
            'Quitar el dialogo del user muerto
130         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
132         .Stats.MinHp = 0
134         .Stats.MinSta = 0
136         .flags.AtacadoPorUser = 0
138         .flags.Envenenado = 0
140         .flags.Ahogandose = 0
142         .flags.Incinerado = 0
144         .flags.incinera = 0
146         .flags.Paraliza = 0
148         .flags.Envenena = 0
150         .flags.Estupidiza = 0
152         .flags.Muerto = 1
            '.flags.SeguroParty = True
            'Call WritePartySafeOn(UserIndex)
        
154         aN = .flags.AtacadoPorNpc
    
156         If aN > 0 Then
158             Npclist(aN).Movement = Npclist(aN).flags.OldMovement
160             Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
162             Npclist(aN).flags.AttackedBy = vbNullString
    
            End If
        
164         aN = .flags.NPCAtacado
    
166         If aN > 0 Then
168             If Npclist(aN).flags.AttackedFirstBy = .name Then
170                 Npclist(aN).flags.AttackedFirstBy = vbNullString
    
                End If
    
            End If
    
172         .flags.AtacadoPorNpc = 0
174         .flags.NPCAtacado = 0
        
            '<<<< Paralisis >>>>
176         If .flags.Paralizado = 1 Then
178             .flags.Paralizado = 0
180             Call WriteParalizeOK(Userindex)
    
            End If
        
            '<<<< Inmovilizado >>>>
182         If .flags.Inmovilizado = 1 Then
184             .flags.Inmovilizado = 0
186             Call WriteInmovilizaOK(Userindex)
    
            End If
        
            '<<< Estupidez >>>
188         If .flags.Estupidez = 1 Then
190             .flags.Estupidez = 0
192             Call WriteDumbNoMore(Userindex)
    
            End If
        
            '<<<< Descansando >>>>
194         If .flags.Descansar Then
196             .flags.Descansar = False
198             Call WriteRestOK(Userindex)
    
            End If
        
            '<<<< Meditando >>>>
200         If .flags.Meditando Then
202             .flags.Meditando = False
204             .Char.FX = 0
206             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
            End If
        
            'If .Familiar.Invocado = 1 Then
            ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso("17", Npclist(.Familiar.Id).Pos.x, Npclist(.Familiar.Id).Pos.Y))
            ' .Familiar.Invocado = 0
            ' Call QuitarNPC(.Familiar.Id)
            ' End If
        
            '<<<< Invisible >>>>
208         If .flags.invisible = 1 Or .flags.Oculto = 1 Then
210             .flags.Oculto = 0
212             .flags.invisible = 0
214             .Counters.TiempoOculto = 0
216             .Counters.Invisibilidad = 0
                'no hace falta encriptar este NOVER
218             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
    
            End If
    
220         If TriggerZonaPelea(Userindex, Userindex) <> eTrigger6.TRIGGER6_PERMITE Then

222             If (.flags.Privilegios And PlayerType.user) <> 0 Then
        
224                 If Not EsNewbie(Userindex) Then
226                     If .flags.PendienteDelSacrificio = 0 Then
                
228                         Call TirarTodo(Userindex)
                    
230                         If .Invent.ArmourEqpObjIndex > 0 Then
232                             If ItemSeCae(.Invent.ArmourEqpObjIndex) Then
234                                 Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

                                End If

                            End If

                            'desequipar arma
236                         If .Invent.WeaponEqpObjIndex > 0 Then
238                             If ItemSeCae(.Invent.WeaponEqpObjIndex) Then
240                                 Call Desequipar(Userindex, .Invent.WeaponEqpSlot)

                                End If

                            End If

                            'desequipar casco
242                         If .Invent.CascoEqpObjIndex > 0 Then
244                             If ItemSeCae(.Invent.CascoEqpObjIndex) Then
246                                 Call Desequipar(Userindex, .Invent.CascoEqpSlot)

                                End If

                            End If

                            'desequipar herramienta
248                         If .Invent.AnilloEqpObjIndex > 0 Then
250                             If ItemSeCae(.Invent.AnilloEqpObjIndex) Then
252                                 Call Desequipar(Userindex, .Invent.AnilloEqpSlot)

                                End If

                            End If

                            'desequipar municiones
254                         If .Invent.MunicionEqpObjIndex > 0 Then
256                             If ItemSeCae(.Invent.MunicionEqpObjIndex) Then
258                                 Call Desequipar(Userindex, .Invent.MunicionEqpSlot)

                                End If

                            End If

                            'desequipar escudo
260                         If .Invent.EscudoEqpObjIndex > 0 Then
262                             If ItemSeCae(.Invent.EscudoEqpObjIndex) Then
264                                 Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

                                End If

                            End If
                    
266                         If .Invent.MagicoObjIndex > 0 Then
268                             If ItemSeCae(.Invent.MagicoObjIndex) Then
270                                 Call Desequipar(Userindex, .Invent.MagicoSlot)

                                End If
                            End If
    
                        Else
                
                            Dim MiObj As obj

272                         MiObj.Amount = 1
274                         MiObj.ObjIndex = PENDIENTE
276                         Call QuitarObjetos(PENDIENTE, 1, Userindex)
278                         Call MakeObj(MiObj, .Pos.Map, .Pos.X, .Pos.Y)
280                         Call WriteConsoleMsg(Userindex, "Has perdido tu pendiente del sacrificio.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else

282                     Call TirarTodosLosItemsNoNewbies(Userindex)

                    End If
    
                End If
    
            End If
        
284         .flags.CarroMineria = 0
            
            ' DESEQUIPA TODOS LOS OBJETOS
            
286         If .Char.Arma_Aura <> "" Then
288             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 1))
290             .Char.Arma_Aura = ""
            End If
        
292         If .Char.Arma_Aura <> "" Then
294             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 1))
296             .Char.Arma_Aura = ""
    
            End If
    
298         If .Char.Body_Aura <> "" Then
300             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 2))
302             .Char.Body_Aura = 0
    
            End If
        
304         If .Char.Escudo_Aura <> "" Then
306             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 3))
308             .Char.Escudo_Aura = 0
    
            End If
        
310         If .Char.Head_Aura <> "" Then
312             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 4))
314             .Char.Head_Aura = 0
    
            End If
        
316         If .Char.Anillo_Aura <> "" Then
318             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 6))
320             .Char.Anillo_Aura = 0
            End If
    
322         If .Char.Otra_Aura <> "" Then
324             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, 0, True, 5))
326             .Char.Otra_Aura = 0
            End If
                            
            'desequipar montura
328         If .flags.Montado > 0 Then
    
330             Call DoMontar(Userindex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
    
            End If
        
            ' << Reseteamos los posibles FX sobre el personaje >>
336         If .Char.loops = INFINITE_LOOPS Then
338             .Char.FX = 0
340             .Char.loops = 0
    
            End If
        
342         .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        
            ' << Restauramos los atributos >>
344         If .flags.TomoPocion = True And .flags.BattleModo = 0 Then
    
346             For i = 1 To 4
348                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
350             Next i
    
352             Call WriteFYA(Userindex)
    
            End If
        
354         .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        
            ' << Restauramos los atributos >>
356         If .flags.TomoPocion = True And .flags.BattleModo = 0 Then
    
358             For i = 1 To 4
360                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
362             Next i
    
364             Call WriteFYA(Userindex)
    
            End If
        
            '<< Cambiamos la apariencia del char >>
366         If .flags.Navegando = 0 Then
368             .Char.Body = iCuerpoMuerto
370             .Char.Head = 0
372             .Char.ShieldAnim = NingunEscudo
374             .Char.WeaponAnim = NingunArma
376             .Char.CascoAnim = NingunCasco
            
                .Char.speeding = VelocidadMuerto
            Else
    
378             If ObjData(.Invent.BarcoObjIndex).Ropaje = iTraje Then
380                 .Char.Body = iRopaBuceoMuerto
382                 .Char.Head = iCabezaMuerto
                Else
384                 .Char.Body = iFragataFantasmal ';)
386                 .Char.Head = 0
                End If
                
                .Char.speeding = ObjData(.Invent.BarcoObjIndex).Velocidad
            End If
            
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        
388         For i = 1 To MAXMASCOTAS
390             If .MascotasIndex(i) > 0 Then
392                 Call MuereNpc(.MascotasIndex(i), 0)
                ' Si estan en agua o zona segura
                Else
394                 .MascotasType(i) = 0
                End If
396         Next i
        
398         .NroMascotas = 0
        
            '<< Actualizamos clientes >>
400         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
            'Call WriteUpdateUserStats(UserIndex)
        
            'If UCase$(MapInfo(.Pos.Map).restrict_mode) = "NEWBIE" Then
            '    .flags.pregunta = 5
            '    Call WritePreguntaBox(UserIndex, "¡Has muerto! ¿Deseas ser resucitado?")
            'End If
        
        End With

        Exit Sub

ErrorHandler:
402     Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)

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
126     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ContarMuerte", Erl)
128     Resume Next
        
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
122     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.ContarPuntoBattle", Erl)
124     Resume Next
        
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
142     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.Tilelibre", Erl)
144     Resume Next
        
End Sub

Sub WarpToLegalPos(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal FX As Boolean = False, Optional ByVal AguaValida As Boolean = False)
        'Santo: Sub para buscar la posición legal mas cercana al objetivo y warpearlo.
        
        On Error GoTo WarpToLegalPos_Err
        

        Dim ALoop As Byte, Find As Boolean, lX As Long, lY As Long

100     Find = False
102     ALoop = 1

104     Do Until Find = True

106         For lX = X - ALoop To X + ALoop
108             For lY = Y - ALoop To Y + ALoop

110                 With MapData(Map, lX, lY)

112                     If .Userindex <= 0 Then
114                         If (.Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And ((.Blocked And FLAG_AGUA) = 0 Or AguaValida) Then
116                             If .TileExit.Map = 0 Then
118                                 If .NpcIndex <= 0 Then
120                                     If .trigger = 0 Then
122                                         Call WarpUserChar(Userindex, Map, lX, lY, FX)
124                                         Find = True
                                            Exit Sub
                                        End If
                                    End If

                                End If

                            End If

                        End If

                    End With

126             Next lY
128         Next lX

130         ALoop = ALoop + 1
        Loop

        
        Exit Sub

WarpToLegalPos_Err:
132     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.WarpToLegalPos", Erl)
134     Resume Next
        
End Sub

Sub WarpUserChar(ByVal Userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
        
        On Error GoTo WarpUserChar_Err
        

        Dim OldMap As Integer

        Dim OldX   As Integer

        Dim OldY   As Integer
    
100     If UserList(Userindex).ComUsu.DestUsu > 0 Then
102         If UserList(UserList(Userindex).ComUsu.DestUsu).flags.UserLogged Then
104             If UserList(UserList(Userindex).ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
106                 Call WriteConsoleMsg(UserList(Userindex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
108                 Call FinComerciarUsu(UserList(Userindex).ComUsu.DestUsu)
                

                End If

            End If

        End If
    
        'Quitar el dialogo
110     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(UserList(Userindex).Char.CharIndex))
    
112     Call WriteRemoveAllDialogs(Userindex)
    
114     OldMap = UserList(Userindex).Pos.Map
116     OldX = UserList(Userindex).Pos.X
118     OldY = UserList(Userindex).Pos.Y
    
120     Call EraseUserChar(Userindex, True)
    
122     If OldMap <> Map Then
        
124         Call WriteChangeMap(Userindex, Map)
            'Call WriteLight(UserIndex, map)
            'Call WriteHora(UserIndex)
        
            ' If MapInfo(OldMap).music_numberLow <> MapInfo(map).music_numberLow Then
            'Call WritePlayMidi(UserIndex, MapInfo(map).music_numberLow, 1)
            'End If
        
126         If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And UserList(Userindex).Stats.ELV < 42 Then
128             Call WriteConsoleMsg(Userindex, "Estas saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.", FontTypeNames.FONTTYPE_WARNING)

            End If
        
130         UserList(Userindex).flags.NecesitaOxigeno = RequiereOxigeno(Map)

132         If UserList(Userindex).flags.NecesitaOxigeno Then
134             Call WriteContadores(Userindex)
136             Call WriteOxigeno(Userindex)

138             If UserList(Userindex).Counters.Oxigeno = 0 Then
140                 UserList(Userindex).flags.Ahogandose = 1

                End If

            End If
            
142         UserList(Userindex).Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
144         UserList(Userindex).flags.Inmunidad = 1

146         If RequiereOxigeno(OldMap) = True And UserList(Userindex).flags.NecesitaOxigeno = False Then  'And UserList(UserIndex).Stats.ELV < 35 Then
        
                'Call WriteConsoleMsg(UserIndex, "Ya no necesitas oxigeno.", FontTypeNames.FONTTYPE_WARNING)
148             Call WriteContadores(Userindex)
150             Call WriteOxigeno(Userindex)

            End If
        
            'Update new Map Users
152         MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
            'Update old Map Users
154         MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

156         If MapInfo(OldMap).NumUsers < 0 Then
158             MapInfo(OldMap).NumUsers = 0

            End If
            
            'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
            Dim nextMap, previousMap As Boolean
            
160         nextMap = distanceToCities(Map).distanceToCity(UserList(Userindex).Hogar) >= 0
162         previousMap = distanceToCities(UserList(Userindex).Pos.Map).distanceToCity(UserList(Userindex).Hogar) >= 0

164         If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            
166         ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
168             UserList(Userindex).flags.lastMap = UserList(Userindex).Pos.Map
            
170         ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el ultimo mapa es 0 ya que no esta en un dungeon)
172             UserList(Userindex).flags.lastMap = 0
            
174         ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
176             UserList(Userindex).flags.lastMap = UserList(Userindex).flags.lastMap

            End If
        

178         If UserList(Userindex).flags.Traveling = 1 Then
180             UserList(Userindex).flags.Traveling = 0
182             UserList(Userindex).Counters.goHome = 0
184             Call WriteConsoleMsg(Userindex, "El viaje ha terminado", FontTypeNames.FONTTYPE_INFOBOLD)
    
            End If

        End If
    
186     UserList(Userindex).Pos.X = X
188     UserList(Userindex).Pos.Y = Y
190     UserList(Userindex).Pos.Map = Map
    
192     If FX Then
194         Call MakeUserChar(True, Map, Userindex, Map, X, Y, 1)
        Else
196         Call MakeUserChar(True, Map, Userindex, Map, X, Y, 0)

        End If
    
198     Call WriteUserCharIndexInServer(Userindex)
    
        'Force a flush, so user index is in there before it's destroyed for teleporting
    
    
        'Seguis invisible al pasar de mapa
200     If (UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1) And (Not UserList(Userindex).flags.AdminInvisible = 1) Then
202         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, True))

        End If
    
        'Reparacion temporal del bug de particulas. 08/07/09 LADDER

        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 71, False))
    
204     If UserList(Userindex).flags.AdminInvisible = 0 Then
206         If FX Then 'FX
208             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_WARP, X, Y))
210             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, FXIDs.FXWARP, 0))
            End If
        Else
212         Call EnviarDatosASlot(Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, True))
        End If
        
214     If UserList(Userindex).NroMascotas > 0 Then Call WarpMascotas(Userindex)
    
216     If MapInfo(Map).zone = "DUNGEON" Then
218         If UserList(Userindex).flags.Montado > 0 Then
220             Call DoMontar(Userindex, ObjData(UserList(Userindex).Invent.MonturaObjIndex), UserList(Userindex).Invent.MonturaSlot)

            End If

        End If
    
        ' Call WarpFamiliar(UserIndex)
        
        Exit Sub

WarpUserChar_Err:
222     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.WarpUserChar", Erl)
224     Resume Next
        
End Sub

Sub WarpFamiliar(ByVal Userindex As Integer)
        
        On Error GoTo WarpFamiliar_Err
        

100     With UserList(Userindex)

102         If .Familiar.Invocado = 1 Then
104             Call QuitarNPC(.Familiar.Id)
                ' If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
106             .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, UserList(Userindex).Pos, False, True)

                'Controlamos que se sumoneo OK
108             If .Familiar.Id = 0 Then
110                 Call WriteConsoleMsg(Userindex, "No hay espacio aquí para tu mascota. Se provoco un ERROR.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

112             Call CargarFamiliar(Userindex)
            Else

                'Call WriteConsoleMsg(UserIndex, "No se permiten familiares en zona segura. " & .Familiar.Nombre & " te esperará afuera.", FontTypeNames.FONTTYPE_INFO)
            End If
    
        End With
            
        
        Exit Sub

WarpFamiliar_Err:
114     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.WarpFamiliar", Erl)
116     Resume Next
        
End Sub

Sub Cerrar_Usuario(ByVal Userindex As Integer)

        On Error GoTo Cerrar_Usuario_Err

        '***************************************************
        'Author: Unknown
        'Last Modification: 16/09/2010
        '16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
        '***************************************************
        Dim isNotVisible As Boolean
        Dim HiddenPirat  As Boolean
    
100     With UserList(Userindex)

102         If .flags.UserLogged And Not .Counters.Saliendo Then
104             .Counters.Saliendo = True
106             .Counters.Salir = IntervaloCerrarConexion
            
108             isNotVisible = (.flags.Oculto Or .flags.invisible)

110             If isNotVisible Then
112                 .flags.invisible = 0
                
114                 If .flags.Oculto Then
                
116                     If .flags.Navegando = 1 Then
                    
118                         If .clase = eClass.Pirat Then
                                ' Pierde la apariencia de fragata fantasmal
120                             .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
        
122                             .Char.ShieldAnim = NingunEscudo
124                             .Char.WeaponAnim = NingunArma
126                             .Char.CascoAnim = NingunCasco
        
128                             Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
130                             Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
132                             HiddenPirat = True

                            End If

                        End If

                    End If
                
134                 .flags.Oculto = 0
                
                    ' Para no repetir mensajes
136                 If Not HiddenPirat Then
138                     Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    End If
                
                    ' Si esta navegando ya esta visible
140                 If .flags.Navegando = 0 Then
142                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    End If

                End If
            
144             If .flags.Traveling = 1 Then
146                 Call WriteConsoleMsg(Userindex, "Se ha cancelado el viaje a casa", FontTypeNames.FONTTYPE_INFO)
148                 .flags.Traveling = 0
150                 .Counters.goHome = 0

                End If
            
152             Call WriteLocaleMsg(Userindex, "203", FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
            
154             If EsGM(Userindex) Or MapInfo(.Pos.Map).Seguro = 1 Then
156                 Call WriteDisconnect(Userindex)
158                 Call CloseSocket(Userindex)
                End If

            End If

        End With

        Exit Sub

Cerrar_Usuario_Err:
160     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.Cerrar_Usuario", Erl)
162     Resume Next

End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal Userindex As Integer)
        
        On Error GoTo CancelExit_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/02/08
        '
        '***************************************************
100     If UserList(Userindex).Counters.Saliendo Then

            ' Is the user still connected?
102         If UserList(Userindex).ConnIDValida Then
104             UserList(Userindex).Counters.Saliendo = False
106             UserList(Userindex).Counters.Salir = 0
108             Call WriteConsoleMsg(Userindex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Else

                'Simply reset
110             If UserList(Userindex).flags.Privilegios = PlayerType.user And MapInfo(UserList(Userindex).Pos.Map).Seguro = 0 Then
112                 UserList(Userindex).Counters.Salir = IntervaloCerrarConexion
                Else
114                 Call WriteConsoleMsg(Userindex, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
116                 Call WriteDisconnect(Userindex)
                
118                 Call CloseSocket(Userindex)

                End If
            
                'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
            End If

        End If

        
        Exit Sub

CancelExit_Err:
120     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.CancelExit", Erl)
122     Resume Next
        
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal Userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
        
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
110     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.CambiarNick", Erl)
112     Resume Next
        
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
124     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.SendUserStatsTxtOFF", Erl)
126     Resume Next
        
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

        On Error Resume Next

        Dim j        As Integer

        Dim CharFile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long

100     CharFile = CharPath & CharName & ".chr"

102     If FileExist(CharFile, vbNormal) Then
104         Call WriteConsoleMsg(sendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
        Else
108         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

Sub VolverCriminal(ByVal Userindex As Integer)
        
        On Error GoTo VolverCriminal_Err
        

        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/06/2006
        'Nacho: Actualiza el tag al cliente
        '**************************************************************
100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub

102     If UserList(Userindex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
   
104         If UserList(Userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)

        End If

106     If UserList(Userindex).Faccion.FuerzasCaos = 1 Then Exit Sub

108     UserList(Userindex).Faccion.Status = 0

110     Call RefreshCharStatus(Userindex)

        
        Exit Sub

VolverCriminal_Err:
112     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.VolverCriminal", Erl)
114     Resume Next
        
End Sub

Sub VolverCiudadano(ByVal Userindex As Integer)
        '**************************************************************
        'Author: Unknown
        'Last Modify Date: 21/06/2006
        'Nacho: Actualiza el tag al cliente.
        '**************************************************************
        
        On Error GoTo VolverCiudadano_Err
        

100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub

102     UserList(Userindex).Faccion.Status = 1
104     Call RefreshCharStatus(Userindex)

        
        Exit Sub

VolverCiudadano_Err:
106     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.VolverCiudadano", Erl)
108     Resume Next
        
End Sub

Public Function getMaxInventorySlots(ByVal Userindex As Integer) As Byte
        '***************************************************
        'Author: Unknown
        'Last Modification: 30/09/2020
        '
        '***************************************************
        
        On Error GoTo getMaxInventorySlots_Err
        

100     If UserList(Userindex).Stats.InventLevel > 0 Then
102         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS + UserList(Userindex).Stats.InventLevel * SLOTS_PER_ROW_INVENTORY
        Else
104         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS

        End If

        
        Exit Function

getMaxInventorySlots_Err:
106     Call RegistrarError(Err.Number, Err.description, "UsUaRiOs.getMaxInventorySlots", Erl)
108     Resume Next
        
End Function

Private Sub WarpMascotas(ByVal Userindex As Integer)

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

        Dim canWarp          As Boolean

        Dim index            As Integer

        Dim iMinHP           As Integer
        
        Dim PetTiempoDeVida  As Integer
    
        Dim MascotaQuitada   As Boolean
        Dim ElementalQuitado As Boolean
        Dim SpawnInvalido    As Boolean

100     canWarp = MapInfo(UserList(Userindex).Pos.Map).Seguro = 0

102     For i = 1 To MAXMASCOTAS
104         index = UserList(Userindex).MascotasIndex(i)
        
106         If index > 0 Then
108             iMinHP = Npclist(index).Stats.MinHp
                PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
            
112             Npclist(index).MaestroUser = 0
            
114             Call QuitarNPC(index)

                If PetTiempoDeVida > 0 Then
120                 Call QuitarMascota(Userindex, index)
122                 ElementalQuitado = True

123             ElseIf Not canWarp Then
124                 UserList(Userindex).MascotasIndex(i) = 0
126                 MascotaQuitada = True
                End If
            
            Else
128             iMinHP = 0
130             PetTiempoDeVida = 0
            End If
        
132         petType = UserList(Userindex).MascotasType(i)
        
134         If petType > 0 And canWarp And UserList(Userindex).flags.MascotasGuardadas = 0 And PetTiempoDeVida = 0 Then
        
                Dim SpawnPos As WorldPos
        
136             SpawnPos.Map = UserList(Userindex).Pos.Map
138             SpawnPos.X = UserList(Userindex).Pos.X + RandomNumber(-3, 3)
140             SpawnPos.Y = UserList(Userindex).Pos.Y + RandomNumber(-3, 3)
        
142             index = SpawnNpc(petType, SpawnPos, False, PetRespawn)
            
                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
144             If index > 0 Then
146                 UserList(Userindex).MascotasIndex(i) = index

                    ' Nos aseguramos de que conserve el hp, si estaba danado
148                 If iMinHP Then Npclist(index).Stats.MinHp = iMinHP
            
150                 Npclist(index).MaestroUser = Userindex
154                 Call FollowAmo(index)
            
                Else
156                 SpawnInvalido = True
                End If

            End If

158     Next i

160     If canWarp Then
162         If SpawnInvalido Then
164             Call WriteConsoleMsg(Userindex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
166         If MascotaQuitada Then
168             Call WriteConsoleMsg(Userindex, "No se permiten mascotas en zona segura. Estas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
            End If
        
170         If ElementalQuitado Then
172             Call WriteConsoleMsg(Userindex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If

End Sub

Function TieneArmaduraCazador(ByVal Userindex As Integer) As Boolean

100     If UserList(Userindex).Invent.ArmourEqpObjIndex > 0 Then
        
102         If ObjData(UserList(Userindex).Invent.ArmourEqpObjIndex).Subtipo = 3 Then ' Aguante hardcodear números :D
104             TieneArmaduraCazador = True
            End If
        
        End If

End Function
