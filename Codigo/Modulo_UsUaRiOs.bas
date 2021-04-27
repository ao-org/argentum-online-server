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

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
        
        On Error GoTo ActStats_Err
        

        Dim DaExp       As Integer

        Dim EraCriminal As Byte
    
100     DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)
    
102     If UserList(AttackerIndex).Stats.ELV < STAT_MAXELV Then
104         UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp

106         If UserList(AttackerIndex).Stats.Exp > MAXEXP Then UserList(AttackerIndex).Stats.Exp = MAXEXP

108         Call WriteUpdateExp(AttackerIndex)
110         Call CheckUserLevel(AttackerIndex)

        End If
    
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
    
112     Call WriteLocaleMsg(AttackerIndex, "184", FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).name)
114     Call WriteLocaleMsg(AttackerIndex, "140", FontTypeNames.FONTTYPE_EXP, DaExp)
          
        'Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
116     Call WriteLocaleMsg(VictimIndex, "185", FontTypeNames.FONTTYPE_FIGHT, UserList(AttackerIndex).name)
    
118     If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
120         EraCriminal = Status(AttackerIndex)
        
122         If EraCriminal = 2 And Status(AttackerIndex) < 2 Then
124             Call RefreshCharStatus(AttackerIndex)
126         ElseIf EraCriminal < 2 And Status(AttackerIndex) = 2 Then
128             Call RefreshCharStatus(AttackerIndex)

            End If

        End If
    
130     Call UserDie(VictimIndex)
        
136     If UserList(AttackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then
            UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1
        End If
        
        Exit Sub

ActStats_Err:
138     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.ActStats", Erl)
140     Resume Next
        
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer, Optional ByVal MedianteHechizo As Boolean)
        
        On Error GoTo RevivirUsuario_Err
        
100     With UserList(UserIndex)

102         .flags.Muerto = 0
104         .Stats.MinHp = .Stats.MaxHp

            ' El comportamiento cambia si usamos el hechizo Resucitar
106         If MedianteHechizo Then
108             .Stats.MinHp = 1
110             .Stats.MinHam = 0
112             .Stats.MinAGU = 0
            
114             Call WriteUpdateHungerAndThirst(UserIndex)
            End If
        
118         Call WriteUpdateHP(UserIndex)
            
120         If .flags.Navegando = 1 Then
                Call EquiparBarco(UserIndex)
            Else

146             .Char.Head = .OrigChar.Head
    
148             If .Invent.CascoEqpObjIndex > 0 Then
150                 .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                End If
    
152             If .Invent.EscudoEqpObjIndex > 0 Then
154                 .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
    
                End If
    
156             If .Invent.WeaponEqpObjIndex > 0 Then
158                 .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
        
160                 If ObjData(.Invent.WeaponEqpObjIndex).CreaGRH <> "" Then
162                     .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH
164                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
    
                    End If
            
                End If
    
166             If .Invent.ArmourEqpObjIndex > 0 Then
168                 .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        
170                 If ObjData(.Invent.ArmourEqpObjIndex).CreaGRH <> "" Then
172                     .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH
174                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))
    
                    End If
    
                Else
176                 Call DarCuerpoDesnudo(UserIndex)
            
                End If
    
178             If .Invent.EscudoEqpObjIndex > 0 Then
180                 .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
    
182                 If ObjData(.Invent.EscudoEqpObjIndex).CreaGRH <> "" Then
184                     .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH
186                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
    
                    End If
            
                End If
    
188             If .Invent.CascoEqpObjIndex > 0 Then
190                 .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
    
192                 If ObjData(.Invent.CascoEqpObjIndex).CreaGRH <> "" Then
194                     .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH
196                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
    
                    End If
            
                End If
    
198             If .Invent.MagicoObjIndex > 0 Then
200                 If ObjData(.Invent.MagicoObjIndex).CreaGRH <> "" Then
202                     .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
204                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
    
                    End If
    
                End If
    
206             If .Invent.NudilloObjIndex > 0 Then
208                 If ObjData(.Invent.NudilloObjIndex).CreaGRH <> "" Then
210                     .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
212                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
    
                    End If
                End If
                
214             If .Invent.DañoMagicoEqpObjIndex > 0 Then
216                 If ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH <> "" Then
218                     .Char.DM_Aura = ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH
220                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.DM_Aura, False, 6))
                    End If
                End If
                
222             If .Invent.ResistenciaEqpObjIndex > 0 Then
224                 If ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH <> "" Then
226                     .Char.RM_Aura = ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH
228                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.RM_Aura, False, 7))
                    End If
                End If
    
            End If
    
            Call ActualizarVelocidadDeUsuario(UserIndex)
232         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        End With
        
        Exit Sub

RevivirUsuario_Err:
236     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.RevivirUsuario", Erl)
238     Resume Next
        
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
    
114     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, Head, Heading, UserList(UserIndex).Char.CharIndex, Arma, Escudo, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, Casco, False, UserList(UserIndex).flags.Navegando))

        
        Exit Sub

ChangeUserChar_Err:
116     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserChar", Erl)
118     Resume Next
        
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, ByVal Desvanecer As Boolean)

        On Error GoTo ErrorHandler

        Dim Error As String
   
100     Error = "1"

102     If UserList(UserIndex).Char.CharIndex = 0 Then Exit Sub
   
104     CharList(UserList(UserIndex).Char.CharIndex) = 0
    
106     If UserList(UserIndex).Char.CharIndex = LastChar Then

108         Do Until CharList(LastChar) > 0
110             LastChar = LastChar - 1

112             If LastChar <= 1 Then Exit Do
            Loop

        End If

114     Error = "2"
    
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
116     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex, Desvanecer))
118     Error = "3"
120     Call QuitarUser(UserIndex, UserList(UserIndex).pos.Map)
122     Error = "4"
    
124     MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = 0
126     Error = "5"
128     UserList(UserIndex).Char.CharIndex = 0
    
130     NumChars = NumChars - 1
132     Error = "6"
        Exit Sub
    
ErrorHandler:
134     Call LogError("Error en EraseUserchar " & Error & " - " & Err.Number & ": " & Err.Description)

End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
        
        On Error GoTo RefreshCharStatus_Err
        

        '*************************************************
        'Author: Tararira
        'Last modified: 6/04/2007
        'Refreshes the status and tag of UserIndex.
        '*************************************************
        Dim klan As String, name As String

100     If UserList(UserIndex).showName Then

102         If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.Desactivado Then

104             If UserList(UserIndex).GuildIndex > 0 Then
106                 klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
108                 klan = " <" & klan & ">"
                End If
            
110             name = UserList(UserIndex).name & klan

            Else
112             name = UserList(UserIndex).NameMimetizado
            End If
            
        End If
    
114     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.Status, name))

        
        Exit Sub

RefreshCharStatus_Err:
116     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.RefreshCharStatus", Erl)
118     Resume Next
        
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, _
                 ByVal sndIndex As Integer, _
                 ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 Optional ByVal appear As Byte = 0)

        On Error GoTo HayError

        Dim CharIndex As Integer

        Dim TempName  As String
    
100     If InMapBounds(Map, x, y) Then
        
102         With UserList(UserIndex)
        
                'If needed make a new character in list
104             If .Char.CharIndex = 0 Then
106                 CharIndex = NextOpenCharIndex
108                 .Char.CharIndex = CharIndex
110                 CharList(CharIndex) = UserIndex
                End If

                'Place character on map if needed
112             If toMap Then MapData(Map, x, y).UserIndex = UserIndex

                'Send make character command to clients
                Dim klan       As String
                Dim clan_nivel As Byte

114             If Not toMap Then
                
116                 If .showName Then
118                     If .flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
120                         If .GuildIndex > 0 Then
                    
122                             klan = modGuilds.GuildName(.GuildIndex)
124                             clan_nivel = modGuilds.NivelDeClan(.GuildIndex)
126                             TempName = .name & " <" & klan & ">"
                    
                            Else
                        
128                             klan = vbNullString
130                             clan_nivel = 0
                            
132                             If .flags.EnConsulta Then
                                
134                                 TempName = .name & " [CONSULTA]"
                                
                                Else
                            
136                                 TempName = .name
                            
                                End If
                            
                            End If
                        Else
138                         TempName = .NameMimetizado
                        End If
                    End If

140                 Call WriteCharacterCreate(sndIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.CharIndex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, TempName, .Faccion.Status, .flags.Privilegios, .Char.ParticulaFx, .Char.Head_Aura, .Char.Arma_Aura, .Char.Body_Aura, .Char.DM_Aura, .Char.RM_Aura, .Char.Otra_Aura, .Char.Escudo_Aura, .Char.speeding, 0, appear, .Grupo.Lider, .GuildIndex, clan_nivel, .Stats.MinHp, .Stats.MaxHp, 0, False, .flags.Navegando)
                                         
                Else
            
                    'Hide the name and clan - set privs as normal user
142                 Call AgregarUser(UserIndex, .pos.Map, appear)
                
                End If
            
            End With
        
        End If

        Exit Sub

HayError:
        
        Dim Desc As String
144         Desc = Err.Description & vbNewLine & _
                    " Usuario: " & UserList(UserIndex).name & vbNewLine & _
                    "Pos: " & Map & "-" & x & "-" & y
            
146     Call RegistrarError(Err.Number, Err.Description, "Usuarios.MakeUserChar", Erl())
        
148     Call CloseSocket(UserIndex)

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
        '17/12/2020 WyroX: Distribución normal de las vidas
        '*************************************************

        On Error GoTo ErrHandler

        Dim Pts              As Integer

        Dim AumentoHIT       As Integer

        Dim AumentoMANA      As Integer

        Dim AumentoSta       As Integer

        Dim AumentoHP        As Integer

        Dim WasNewbie        As Boolean

        Dim Promedio         As Double
        
        Dim PromedioObjetivo As Double
        
        Dim PromedioUser     As Double

        Dim aux              As Integer
    
        Dim PasoDeNivel      As Boolean
        Dim experienceToLevelUp As Long

        ' Randomizo las vidas
100     Randomize Time
    
102     With UserList(UserIndex)

104         WasNewbie = EsNewbie(UserIndex)
            experienceToLevelUp = ExpLevelUp(.Stats.ELV)
        
106         Do While .Stats.Exp >= experienceToLevelUp And .Stats.ELV < STAT_MAXELV
            
                'Store it!
                'Call Statistics.UserLevelUp(UserIndex)

108             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 106, 0))
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .pos.x, .pos.y))
112             Call WriteLocaleMsg(UserIndex, "186", FontTypeNames.FONTTYPE_INFO)
            
114             .Stats.Exp = .Stats.Exp - experienceToLevelUp
                
130             Pts = Pts + 5
            
                ' Calculo subida de vida by WyroX
                ' Obtengo el promedio según clase y constitución
132             PromedioObjetivo = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
                ' Obtengo el promedio actual del user
134             PromedioUser = CalcularPromedioVida(UserIndex)
                ' Lo modifico para compensar si está muy bajo o muy alto
136             Promedio = PromedioObjetivo + (PromedioObjetivo - PromedioUser) * DesbalancePromedioVidas
                ' Obtengo un entero al azar con más tendencia al promedio
138             AumentoHP = RandomIntBiased(PromedioObjetivo - RangoVidas, PromedioObjetivo + RangoVidas, Promedio, InfluenciaPromedioVidas)

                ' WyroX: Aumento del resto de stats
140             AumentoSta = ModClase(.clase).AumentoSta
142             AumentoMANA = ModClase(.clase).MultMana * .Stats.UserAtributos(eAtributos.Inteligencia)
144             AumentoHIT = IIf(.Stats.ELV < 36, ModClase(.clase).HitPre36, ModClase(.clase).HitPost36)

146             .Stats.ELV = .Stats.ELV + 1
147             experienceToLevelUp = ExpLevelUp(.Stats.ELV)
                
                'Actualizamos HitPoints
148             .Stats.MaxHp = .Stats.MaxHp + AumentoHP

150             If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                'Actualizamos Stamina
152             .Stats.MaxSta = .Stats.MaxSta + AumentoSta

154             If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
                'Actualizamos Mana
156             .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

158             If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

                'Actualizamos Golpe Máximo
160             .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            
                'Actualizamos Golpe Mínimo
162             .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
        
                'Notificamos al user
164             If AumentoHP > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
166                 Call WriteLocaleMsg(UserIndex, "197", FontTypeNames.FONTTYPE_INFO, AumentoHP)

                End If

168             If AumentoSta > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
170                 Call WriteLocaleMsg(UserIndex, "198", FontTypeNames.FONTTYPE_INFO, AumentoSta)

                End If

172             If AumentoMANA > 0 Then
174                 Call WriteLocaleMsg(UserIndex, "199", FontTypeNames.FONTTYPE_INFO, AumentoMANA)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
                End If

176             If AumentoHIT > 0 Then
178                 Call WriteLocaleMsg(UserIndex, "200", FontTypeNames.FONTTYPE_INFO, AumentoHIT)

                    'Call WriteConsoleMsg(UserIndex, "Tu golpe aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                End If

180             PasoDeNivel = True
             
                ' Call LogDesarrollo(.name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)

182             .Stats.MinHp = .Stats.MaxHp
            
                ' Call UpdateUserInv(True, UserIndex, 0)
            
184             If OroPorNivel > 0 Then
186                 If EsNewbie(UserIndex) Then
                        Dim OroRecompenza As Long
    
188                     OroRecompenza = OroPorNivel * .Stats.ELV * OroMult
190                     .Stats.GLD = .Stats.GLD + OroRecompenza
                        'Call WriteConsoleMsg(UserIndex, "Has ganado " & OroRecompenza & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
192                     Call WriteLocaleMsg(UserIndex, "29", FontTypeNames.FONTTYPE_INFO, PonerPuntos(OroRecompenza))
                    End If
                End If
            
194             If Not EsNewbie(UserIndex) And WasNewbie Then
        
196                 Call QuitarNewbieObj(UserIndex)
            
                End If
        
            Loop
        
198         If PasoDeNivel Then
200             If .Stats.ELV >= STAT_MAXELV Then .Stats.Exp = 0
        
206             Call UpdateUserInv(True, UserIndex, 0)
                'Call CheckearRecompesas(UserIndex, 3)
208             Call WriteUpdateUserStats(UserIndex)
            
210             If Pts > 0 Then
                
212                 .Stats.SkillPts = .Stats.SkillPts + Pts
214                 Call WriteLevelUp(UserIndex, .Stats.SkillPts)
216                 Call WriteLocaleMsg(UserIndex, "187", FontTypeNames.FONTTYPE_INFO, Pts)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
                End If
                
218             If .Stats.ELV >= MapInfo(.pos.Map).MaxLevel And Not EsGM(UserIndex) Then
220                 If MapInfo(.pos.Map).Salida.Map <> 0 Then
222                     Call WriteConsoleMsg(UserIndex, "Tu nivel no te permite seguir en el mapa.", FontTypeNames.FONTTYPE_INFO)
224                     Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
                    End If
                End If

            End If
    
        End With
    
        Exit Sub

ErrHandler:
226     Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description)

End Sub

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading) As Boolean
    ' 20/01/2021 - WyroX: Lo convierto a función y saco los WritePosUpdate, ahora están en el paquete

    On Error GoTo MoveUserChar_Err

    Dim nPos         As WorldPos
    Dim nPosOriginal As WorldPos
    Dim nPosMuerto   As WorldPos
    Dim IndexMover As Integer
    Dim OppositeHeading As eHeading

    With UserList(UserIndex)

        nPos = .pos
        Call HeadtoPos(nHeading, nPos)

        If Not LegalWalk(.pos.Map, nPos.x, nPos.y, nHeading, .flags.Navegando = 1, .flags.Navegando = 0, .flags.Montado) Then
            Exit Function
        End If

        If .Accion.AccionPendiente = True Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, .Accion.Particula, 1, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 1, Accion_Barra.CancelarAccion))
            .Accion.AccionPendiente = False
            .Accion.Particula = 0
            .Accion.TipoAccion = Accion_Barra.CancelarAccion
            .Accion.HechizoPendiente = 0
            .Accion.RunaObj = 0
            .Accion.ObjSlot = 0
            .Accion.AccionPendiente = False
        End If

        If .flags.Muerto = 0 Then
            If MapData(nPos.Map, nPos.x, nPos.y).TileExit.Map <> 0 And .Counters.TiempoDeMapeo > 0 Then
                Call WriteConsoleMsg(UserIndex, "Estás en combate, debes aguardar " & .Counters.TiempoDeMapeo & " segundo(s) para escapar...", FontTypeNames.FONTTYPE_INFOBOLD)
                Exit Function
            End If
        End If

        'Si no estoy solo en el mapa...
        If MapInfo(.pos.Map).NumUsers > 1 Then
            ' Intercambia posición si hay un casper o gm invisible
            IndexMover = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
            
            If IndexMover <> 0 Then
                ' Sólo puedo patear caspers/gms invisibles si no es él un gm invisible
                If .flags.AdminInvisible <> 0 Then Exit Function

                Call WritePosUpdate(IndexMover)
                OppositeHeading = InvertHeading(nHeading)
                Call HeadtoPos(OppositeHeading, UserList(IndexMover).pos)
                
                ' Si es un admin invisible, no se avisa a los demas clientes
                If UserList(IndexMover).flags.AdminInvisible = 0 Then
                    Call SendData(SendTarget.ToPCAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.CharIndex, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y))
                Else
                    Call SendData(SendTarget.ToAdminAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.CharIndex, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y))
                End If
                Call WriteForceCharMove(IndexMover, OppositeHeading)
                
                'Update map and char
                UserList(IndexMover).Char.Heading = OppositeHeading
                MapData(UserList(IndexMover).pos.Map, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y).UserIndex = IndexMover
                
                'Actualizamos las areas de ser necesario
                Call ModAreas.CheckUpdateNeededUser(IndexMover, OppositeHeading, 0)
            End If

            If .flags.AdminInvisible = 0 Then
                Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.x, nPos.y))
            Else
                Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.x, nPos.y))
            End If
        End If
        
        'Update map and user pos
        If MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex Then
            MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0
        End If

        .pos = nPos
        .Char.Heading = nHeading
        MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)

        If .Counters.Trabajando Then
            Call WriteMacroTrabajoToggle(UserIndex, False)
        End If

        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    
    End With
    
    MoveUserChar = True
    
    Exit Function
    
MoveUserChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.MoveUserChar", Erl)
    Resume Next
        
End Function

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
        
        On Error GoTo InvertHeading_Err
    
        

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

        
        Exit Function

InvertHeading_Err:
116     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.InvertHeading", Erl)

        
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal slot As Byte, ByRef Object As UserOBJ)
        
        On Error GoTo ChangeUserInv_Err
        
100     UserList(UserIndex).Invent.Object(slot) = Object
102     Call WriteChangeInventorySlot(UserIndex, slot)

        
        Exit Sub

ChangeUserInv_Err:
104     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserInv", Erl)
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
112     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.NextOpenCharIndex", Erl)
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
110     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.NextOpenUser", Erl)
112     Resume Next
        
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserStatsTxt_Err
        

        Dim GuildI As Integer

100     Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & ExpLevelUp(UserList(UserIndex).Stats.ELV), FontTypeNames.FONTTYPE_INFO)
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

148     Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).pos.x & "," & UserList(UserIndex).pos.y & " en mapa " & UserList(UserIndex).pos.Map, FontTypeNames.FONTTYPE_INFO)
150     Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma), FontTypeNames.FONTTYPE_INFO)
152     Call WriteConsoleMsg(sendIndex, "Veces que Moriste: " & UserList(UserIndex).flags.VecesQueMoriste, FontTypeNames.FONTTYPE_INFO)

        Exit Sub

SendUserStatsTxt_Err:
154     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserStatsTxt", Erl)
156     Resume Next
        
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
104         Call WriteConsoleMsg(sendIndex, "Ciudadanos Matados: " & .Faccion.ciudadanosMatados & " Criminales Matados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)

112         If .GuildIndex > 0 Then
114             Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)

            End If

116         Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & .Stats.GLD, FontTypeNames.FONTTYPE_INFO)
118         Call WriteConsoleMsg(sendIndex, "Oro en banco: " & .Stats.Banco, FontTypeNames.FONTTYPE_INFO)
    
120         Call WriteConsoleMsg(sendIndex, "Cuenta: " & .Cuenta, FontTypeNames.FONTTYPE_INFO)
    
        End With

        Exit Sub

SendUserMiniStatsTxt_Err:
126     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)
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
136     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserMiniStatsTxtFromChar", Erl)
138     Resume Next
        
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserInvTxt_Err
    
        

        

        Dim j As Long
    
100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
104     For j = 1 To UserList(UserIndex).CurrentInventorySlots

106         If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)

            End If

110     Next j

        
        Exit Sub

SendUserInvTxt_Err:
112     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserInvTxt", Erl)

        
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
        
        On Error GoTo SendUserInvTxtFromChar_Err
    
        

        

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
    
        
        Exit Sub

SendUserInvTxtFromChar_Err:
124     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserInvTxtFromChar", Erl)

        
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserSkillsTxt_Err
    
        

        

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)

102     For j = 1 To NUMSKILLS
104         Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
        Next
106     Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

        
        Exit Sub

SendUserSkillsTxt_Err:
108     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserSkillsTxt", Erl)

        
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
112     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.DameUserIndex", Erl)
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
114     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.DameUserIndexConNombre", Erl)
116     Resume Next
        
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo NPCAtacado_Err
        
        ' WyroX: El usuario pierde la protección
100     UserList(UserIndex).Counters.TiempoDeInmunidad = 0
102     UserList(UserIndex).flags.Inmunidad = 0

        'Guardamos el usuario que ataco el npc.
104     If NpcList(NpcIndex).Movement <> Estatico And NpcList(NpcIndex).flags.AttackedFirstBy = vbNullString Then
106         NpcList(NpcIndex).Target = UserIndex
            NpcList(NpcIndex).Hostile = 1
110         NpcList(NpcIndex).flags.AttackedBy = UserList(UserIndex).name
        End If

        'Npc que estabas atacando.
        Dim LastNpcHit As Integer

112     LastNpcHit = UserList(UserIndex).flags.NPCAtacado
        'Guarda el NPC que estas atacando ahora.
114     UserList(UserIndex).flags.NPCAtacado = NpcIndex

134     If NpcList(NpcIndex).flags.Faccion = Armada And Status(UserIndex) = e_Facciones.Ciudadano Then
            Call VolverCriminal(UserIndex)
        End If
        
140     If NpcList(NpcIndex).MaestroUser > 0 And NpcList(NpcIndex).MaestroUser <> UserIndex Then
144         Call AllMascotasAtacanUser(UserIndex, NpcList(NpcIndex).MaestroUser)
        End If

146     Call AllMascotasAtacanNPC(NpcIndex, UserIndex)
        
        Exit Sub

NPCAtacado_Err:
148     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.NPCAtacado", Erl)
150     Resume Next
        
End Sub

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
        On Error GoTo SubirSkill_Err

        Dim Lvl As Integer, maxPermitido As Integer
            Lvl = UserList(UserIndex).Stats.ELV

100     If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

        ' Se suben 5 skills cada dos niveles como máximo.
        If (Lvl Mod 2 = 0) Then ' El level es numero par
          maxPermitido = (Lvl \ 2) * 5
        Else ' El level es numero impar
          ' Esta cuenta signifca, que si el nivel anterior terminaba en 5 ahora
          ' suma dos puntos mas, sino 3. Lo de siempre.
          maxPermitido = (Lvl \ 2) * 5 + 3 - (((((Lvl - 1) \ 2) * 5) Mod 10) \ 5)
        End If

        If UserList(UserIndex).Stats.UserSkills(Skill) >= maxPermitido Then Exit Sub

102     If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then

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

142             BonusExp = 50 * ExpMult
        
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
162     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SubirSkill", Erl)
164     Resume Next
        
End Sub

Public Sub SubirSkillDeArmaActual(ByVal UserIndex As Integer)
    On Error GoTo SubirSkillDeArmaActual_Err

    With UserList(UserIndex)

        If .Invent.WeaponEqpObjIndex > 0 Then
            ' Arma con proyectiles, subimos armas a distancia
            If ObjData(.Invent.WeaponEqpObjIndex).Proyectil Then
                Call SubirSkill(UserIndex, eSkill.Proyectiles)

            ' Sino, subimos combate con armas
            Else
                Call SubirSkill(UserIndex, eSkill.Armas)
            End If

        ' Si no está usando un arma, subimos combate sin armas
        Else
            Call SubirSkill(UserIndex, eSkill.Wrestling)
        End If

    End With

    Exit Sub

SubirSkillDeArmaActual_Err:
        Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SubirSkillDeArmaActual", Erl)
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
    
100     With UserList(UserIndex)
    
            'Sonido
102         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(IIf(.genero = eGenero.Hombre, e_SoundIndex.MUERTE_HOMBRE, e_SoundIndex.MUERTE_MUJER), .pos.x, .pos.y))
        
            'Quitar el dialogo del user muerto
104         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
106         .Stats.MinHp = 0
108         .Stats.MinSta = 0
110         .flags.AtacadoPorUser = 0

118         .flags.incinera = 0
120         .flags.Paraliza = 0
122         .flags.Envenena = 0
124         .flags.Estupidiza = 0
126         .flags.Muerto = 1
            .flags.Ahogandose = 0
            
128         Call WriteUpdateHP(UserIndex)
130         Call WriteUpdateSta(UserIndex)
        
132         aN = .flags.AtacadoPorNpc
    
134         If aN > 0 Then
136             NpcList(aN).Movement = NpcList(aN).flags.OldMovement
138             NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
140             NpcList(aN).flags.AttackedBy = vbNullString
                NpcList(aN).Target = 0
            End If
        
142         aN = .flags.NPCAtacado
    
144         If aN > 0 Then
146             If NpcList(aN).flags.AttackedFirstBy = .name Then
148                 NpcList(aN).flags.AttackedFirstBy = vbNullString
                End If
            End If
    
150         .flags.AtacadoPorNpc = 0
152         .flags.NPCAtacado = 0
    
198         If MapData(.pos.Map, .pos.x, .pos.y).trigger <> eTrigger.ZONAPELEA Then

200             If (.flags.Privilegios And PlayerType.user) <> 0 Then

202                 If .flags.PendienteDelSacrificio = 0 Then
                
204                     Call TirarTodosLosItems(UserIndex)
    
                    Else
                
                        Dim MiObj As obj

206                     MiObj.amount = 1
208                     MiObj.ObjIndex = PENDIENTE
210                     Call QuitarObjetos(PENDIENTE, 1, UserIndex)
212                     Call MakeObj(MiObj, .pos.Map, .pos.x, .pos.y)
214                     Call WriteConsoleMsg(UserIndex, "Has perdido tu pendiente del sacrificio.", FontTypeNames.FONTTYPE_INFO)

                    End If
    
                End If
    
            End If
        
216         .flags.CarroMineria = 0
   
            'desequipar montura
218         If .flags.Montado > 0 Then
220             Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
            End If
        
            ' << Reseteamos los posibles FX sobre el personaje >>
222         If .Char.loops = INFINITE_LOOPS Then
224             .Char.FX = 0
226             .Char.loops = 0
    
            End If
        
228         .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        
            ' << Restauramos los atributos >>
230         If .flags.TomoPocion Then
    
232             For i = 1 To 4
234                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
236             Next i
    
238             Call WriteFYA(UserIndex)
    
            End If
        
            '<< Cambiamos la apariencia del char >>
252         If .flags.Navegando = 0 Then
254             .Char.Body = iCuerpoMuerto
256             .Char.Head = 0
258             .Char.ShieldAnim = NingunEscudo
260             .Char.WeaponAnim = NingunArma
262             .Char.CascoAnim = NingunCasco
            Else
                Call EquiparBarco(UserIndex)
            End If
            
            Call ActualizarVelocidadDeUsuario(UserIndex)
            Call LimpiarEstadosAlterados(UserIndex)
        
280         For i = 1 To MAXMASCOTAS
282             If .MascotasIndex(i) > 0 Then
284                 Call MuereNpc(.MascotasIndex(i), 0)
                End If
286         Next i
        
288         .NroMascotas = 0
        
            '<< Actualizamos clientes >>
290         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

            If MapInfo(.pos.Map).Seguro = 0 Then
                Call WriteConsoleMsg(UserIndex, "Escribe /HOGAR si deseas regresar rápido a tu hogar.", FontTypeNames.FONTTYPE_New_Naranja)
            End If
            
            If .flags.EnReto Then
                Call MuereEnReto(UserIndex)
            End If

        End With

        Exit Sub

ErrorHandler:
292     Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
        On Error GoTo ContarMuerte_Err


        If EsNewbie(Muerto) Then Exit Sub
        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        If Abs(CInt(UserList(Muerto).Stats.ELV) - CInt(UserList(Atacante).Stats.ELV)) > 14 Then Exit Sub
        If Status(Muerto) = 0 Or Status(Muerto) = 2 Then
            If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
                UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name

                If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then
                    UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
                End If
            End If

        ElseIf Status(Muerto) = 1 Or Status(Muerto) = 3 Then

            If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
                UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name

                If UserList(Atacante).Faccion.ciudadanosMatados < MAXUSERMATADOS Then
                    UserList(Atacante).Faccion.ciudadanosMatados = UserList(Atacante).Faccion.ciudadanosMatados + 1
                End If

            End If

        End If

        Exit Sub

ContarMuerte_Err:
        Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.ContarMuerte", Erl)
        Resume Next

End Sub

Sub Tilelibre(ByRef pos As WorldPos, ByRef nPos As WorldPos, ByRef obj As obj, ByRef Agua As Boolean, ByRef Tierra As Boolean, Optional ByVal InitialPos As Boolean = True)

        
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
102     nPos.Map = pos.Map

104     Do While Not LegalPos(pos.Map, nPos.x, nPos.y, Agua, Tierra) Or hayobj
        
106         If LoopC > 15 Then
108             Notfound = True
                Exit Do

            End If
        
110         For tY = pos.y - LoopC To pos.y + LoopC
112             For tX = pos.x - LoopC To pos.x + LoopC
            
114                 If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                        'We continue if: a - the item is different from 0 and the dropped item or b - the Amount dropped + Amount in map exceeds MAX_INVENTORY_OBJS
116                     hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)

118                     If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.amount + obj.amount > MAX_INVENTORY_OBJS)

120                     If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 And (InitialPos Or (tX <> pos.x And tY <> pos.y)) Then
122                         nPos.x = tX
124                         nPos.y = tY
126                         tX = pos.x + LoopC
128                         tY = pos.y + LoopC

                        End If

                    End If
            
130             Next tX
132         Next tY
        
134         LoopC = LoopC + 1
        
        Loop
    
136     If Notfound = True Then
138         nPos.x = 0
140         nPos.y = 0

        End If

        
        Exit Sub

Tilelibre_Err:
142     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.Tilelibre", Erl)
144     Resume Next
        
End Sub

Sub WarpToLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte, Optional ByVal FX As Boolean = False, Optional ByVal AguaValida As Boolean = False)
        'Santo: Sub para buscar la posición legal mas cercana al objetivo y warpearlo.
        
        On Error GoTo WarpToLegalPos_Err
        

        Dim ALoop As Byte, Find As Boolean, lX As Long, lY As Long

100     Find = False
102     ALoop = 0

104     Do Until Find = True

106         For lX = x - ALoop To x + ALoop
108             For lY = y - ALoop To y + ALoop

110                 With MapData(Map, lX, lY)

112                     If .UserIndex <= 0 Then
                            ' No podemos transportarnos a bloqueos totales
114                         If (.Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES And ((.Blocked And FLAG_AGUA) = 0 Or AguaValida) Then

116                             If .TileExit.Map = 0 Then
118                                 If .NpcIndex <= 0 Then
                                        ' A partir del 50 empiezan las casas privadas, ahí no se puede transportar
120                                     If .trigger < 50 Then
122                                         Call WarpUserChar(UserIndex, Map, lX, lY, FX)
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
132     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.WarpToLegalPos", Erl)
134     Resume Next
        
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal FX As Boolean = False)
        
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
    
114     OldMap = UserList(UserIndex).pos.Map
116     OldX = UserList(UserIndex).pos.x
118     OldY = UserList(UserIndex).pos.y
    
120     Call EraseUserChar(UserIndex, True)
    
122     If OldMap <> Map Then
124         Call WriteChangeMap(UserIndex, Map)
            'Call WriteLight(UserIndex, map)
            'Call WriteHora(UserIndex)
        
            ' If MapInfo(OldMap).music_numberLow <> MapInfo(map).music_numberLow Then
            'Call WritePlayMidi(UserIndex, MapInfo(map).music_numberLow, 1)
            'End If
        
126         If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And UserList(UserIndex).Stats.ELV < 42 Then
128             Call WriteConsoleMsg(UserIndex, "Estás saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.", FontTypeNames.FONTTYPE_WARNING)

            End If
        
130         UserList(UserIndex).flags.NecesitaOxigeno = RequiereOxigeno(Map)

132         If UserList(UserIndex).flags.NecesitaOxigeno Then
134             Call WriteContadores(UserIndex)
136             Call WriteOxigeno(UserIndex)

138             If UserList(UserIndex).Counters.Oxigeno = 0 Then
140                 UserList(UserIndex).flags.Ahogandose = 1

                End If

            End If
            
142         UserList(UserIndex).Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
144         UserList(UserIndex).flags.Inmunidad = 1

146         If RequiereOxigeno(OldMap) = True And UserList(UserIndex).flags.NecesitaOxigeno = False Then  'And UserList(UserIndex).Stats.ELV < 35 Then
        
                'Call WriteConsoleMsg(UserIndex, "Ya no necesitas oxigeno.", FontTypeNames.FONTTYPE_WARNING)
148             Call WriteContadores(UserIndex)
150             Call WriteOxigeno(UserIndex)

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
            
160         nextMap = distanceToCities(Map).distanceToCity(UserList(UserIndex).Hogar) >= 0
162         previousMap = distanceToCities(UserList(UserIndex).pos.Map).distanceToCity(UserList(UserIndex).Hogar) >= 0

164         If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            
166         ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
168             UserList(UserIndex).flags.lastMap = UserList(UserIndex).pos.Map
            
170         ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el ultimo mapa es 0 ya que no esta en un dungeon)
172             UserList(UserIndex).flags.lastMap = 0
            
174         ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
176             UserList(UserIndex).flags.lastMap = UserList(UserIndex).flags.lastMap

            End If

178         If UserList(UserIndex).flags.Traveling = 1 Then
180             UserList(UserIndex).flags.Traveling = 0
182             UserList(UserIndex).Counters.goHome = 0
184             Call WriteConsoleMsg(UserIndex, "El viaje ha terminado.", FontTypeNames.FONTTYPE_INFOBOLD)
    
            End If

        End If
    
186     UserList(UserIndex).pos.x = x
188     UserList(UserIndex).pos.y = y
190     UserList(UserIndex).pos.Map = Map

        If UserList(UserIndex).Grupo.EnGrupo = True Then
            Call CompartirUbicacion(UserIndex)
        End If
    
192     If FX Then
194         Call MakeUserChar(True, Map, UserIndex, Map, x, y, 1)
        Else
196         Call MakeUserChar(True, Map, UserIndex, Map, x, y, 0)

        End If
    
198     Call WriteUserCharIndexInServer(UserIndex)
    
        'Seguis invisible al pasar de mapa
200     If (UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1) And (Not UserList(UserIndex).flags.AdminInvisible = 1) Then
            ' Si el mapa lo permite
202         If MapInfo(Map).SinInviOcul Then
204             UserList(UserIndex).flags.invisible = 0
206             UserList(UserIndex).flags.Oculto = 0
208             UserList(UserIndex).Counters.TiempoOculto = 0
210             Call WriteConsoleMsg(UserIndex, "Una fuerza divina que vigila esta zona te ha vuelto visible.", FontTypeNames.FONTTYPE_INFO)
            Else
212             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
            End If

        End If
    
        'Reparacion temporal del bug de particulas. 08/07/09 LADDER

        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 71, False))
    
214     If UserList(UserIndex).flags.AdminInvisible = 0 Then
216         If FX Then 'FX
218             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, y))
220             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXIDs.FXWARP, 0))
            End If
        Else
222         Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
        End If
        
224     If UserList(UserIndex).NroMascotas > 0 Then Call WarpMascotas(UserIndex)
    
226     If MapInfo(Map).zone = "DUNGEON" Then
228         If UserList(UserIndex).flags.Montado > 0 Then
230             Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

            End If

        End If
    
        ' Call WarpFamiliar(UserIndex)
        
        Exit Sub

WarpUserChar_Err:
232     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.WarpUserChar", Erl)
234     Resume Next
        
End Sub

Sub WarpFamiliar(ByVal UserIndex As Integer)
        
        On Error GoTo WarpFamiliar_Err
        

100     With UserList(UserIndex)

102         If .Familiar.Invocado = 1 Then
104             Call QuitarNPC(.Familiar.Id)
                ' If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
106             .Familiar.Id = SpawnNpc(.Familiar.NpcIndex, UserList(UserIndex).pos, False, True)

                'Controlamos que se sumoneo OK
108             If .Familiar.Id = 0 Then
110                 Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

112             Call CargarFamiliar(UserIndex)
            Else

                'Call WriteConsoleMsg(UserIndex, "No se permiten familiares en zona segura. " & .Familiar.Nombre & " te esperará afuera.", FontTypeNames.FONTTYPE_INFO)
            End If
    
        End With
            
        
        Exit Sub

WarpFamiliar_Err:
114     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.WarpFamiliar", Erl)
116     Resume Next
        
End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer)

        On Error GoTo Cerrar_Usuario_Err
    
100     With UserList(UserIndex)

102         If .flags.UserLogged And Not .Counters.Saliendo Then
104             .Counters.Saliendo = True
106             .Counters.Salir = IntervaloCerrarConexion
            
144             If .flags.Traveling = 1 Then
146                 Call WriteConsoleMsg(UserIndex, "Se ha cancelado el viaje a casa", FontTypeNames.FONTTYPE_INFO)
148                 .flags.Traveling = 0
150                 .Counters.goHome = 0
                End If
            
152             Call WriteLocaleMsg(UserIndex, "203", FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
            
154             If EsGM(UserIndex) Or MapInfo(.pos.Map).Seguro = 1 Then
156                 Call WriteDisconnect(UserIndex)
158                 Call CloseSocket(UserIndex)
                End If

            End If

        End With

        Exit Sub

Cerrar_Usuario_Err:
160     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.Cerrar_Usuario", Erl)
162     Resume Next

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
100     If UserList(UserIndex).Counters.Saliendo And UserList(UserIndex).ConnID <> -1 Then

            ' Is the user still connected?
102         If UserList(UserIndex).ConnIDValida Then
104             UserList(UserIndex).Counters.Saliendo = False
106             UserList(UserIndex).Counters.Salir = 0
108             Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Else

                'Simply reset
110             If UserList(UserIndex).flags.Privilegios = PlayerType.user And MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 Then
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
120     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.CancelExit", Erl)
122     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.CambiarNick", Erl)
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
124     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserStatsTxtOFF", Erl)
126     Resume Next
        
End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
        
        On Error GoTo SendUserOROTxtFromChar_Err
    
        

        

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

        
        Exit Sub

SendUserOROTxtFromChar_Err:
110     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.SendUserOROTxtFromChar", Erl)

        
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
        
    On Error GoTo VolverCriminal_Err
        

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente
    '**************************************************************
        
    With UserList(UserIndex)
        
100     If MapData(.pos.Map, .pos.x, .pos.y).trigger = 6 Then Exit Sub

102     If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
   
104         If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)

        End If

106     If .Faccion.FuerzasCaos = 1 Then Exit Sub

108     .Faccion.Status = 0
        
        If MapInfo(.pos.Map).NoPKs And Not EsGM(UserIndex) And MapInfo(.pos.Map).Salida.Map <> 0 Then
            Call WriteConsoleMsg(UserIndex, "En este mapa no se admiten criminales.", FontTypeNames.FONTTYPE_INFO)
            Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        Else
            Call RefreshCharStatus(UserIndex)
        End If

    End With
        
    Exit Sub

VolverCriminal_Err:
112     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.VolverCriminal", Erl)
114     Resume Next
        
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente.
    '**************************************************************
        
    On Error GoTo VolverCiudadano_Err
        
    With UserList(UserIndex)

100     If MapData(.pos.Map, .pos.x, .pos.y).trigger = 6 Then Exit Sub

102     .Faccion.Status = 1

        If MapInfo(.pos.Map).NoCiudadanos And Not EsGM(UserIndex) And MapInfo(.pos.Map).Salida.Map <> 0 Then
            Call WriteConsoleMsg(UserIndex, "En este mapa no se admiten ciudadanos.", FontTypeNames.FONTTYPE_INFO)
            Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        Else
            Call RefreshCharStatus(UserIndex)
        End If

    End With
        
    Exit Sub

VolverCiudadano_Err:
106     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.VolverCiudadano", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.getMaxInventorySlots", Erl)
108     Resume Next
        
End Function

Private Sub WarpMascotas(ByVal UserIndex As Integer)
        
        On Error GoTo WarpMascotas_Err
    
        

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

        Dim canWarp          As Boolean

        Dim index            As Integer

        Dim iMinHP           As Integer
        
        Dim PetTiempoDeVida  As Integer
    
        Dim MascotaQuitada   As Boolean
        Dim ElementalQuitado As Boolean
        Dim SpawnInvalido    As Boolean

100     canWarp = MapInfo(UserList(UserIndex).pos.Map).Seguro = 0

102     For i = 1 To MAXMASCOTAS
104         index = UserList(UserIndex).MascotasIndex(i)
        
106         If index > 0 Then
108             iMinHP = NpcList(index).Stats.MinHp
110             PetTiempoDeVida = NpcList(index).Contadores.TiempoExistencia
            
112             NpcList(index).MaestroUser = 0
            
114             Call QuitarNPC(index)

116             If PetTiempoDeVida > 0 Then
118                 Call QuitarMascota(UserIndex, index)
120                 ElementalQuitado = True

122             ElseIf Not canWarp Then
124                 UserList(UserIndex).MascotasIndex(i) = 0
126                 MascotaQuitada = True
                End If
            
            Else
128             iMinHP = 0
130             PetTiempoDeVida = 0
            End If
        
132         petType = UserList(UserIndex).MascotasType(i)
        
134         If petType > 0 And canWarp And UserList(UserIndex).flags.MascotasGuardadas = 0 And PetTiempoDeVida = 0 Then
        
                Dim SpawnPos As WorldPos
        
136             SpawnPos.Map = UserList(UserIndex).pos.Map
138             SpawnPos.x = UserList(UserIndex).pos.x + RandomNumber(-3, 3)
140             SpawnPos.y = UserList(UserIndex).pos.y + RandomNumber(-3, 3)
        
142             index = SpawnNpc(petType, SpawnPos, False, False, False, UserIndex)
            
                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
144             If index > 0 Then
146                 UserList(UserIndex).MascotasIndex(i) = index

                    ' Nos aseguramos de que conserve el hp, si estaba danado
148                 If iMinHP Then NpcList(index).Stats.MinHp = iMinHP

152                 Call FollowAmo(index)
            
                Else
154                 SpawnInvalido = True
                End If

            End If

156     Next i

158     If Not canWarp And MascotaQuitada Then
160         Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Estas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)

162     ElseIf SpawnInvalido Then
164         Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)

166     ElseIf ElementalQuitado Then
168         Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
        End If

        
        Exit Sub

WarpMascotas_Err:
170     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.WarpMascotas", Erl)

        
End Sub

Function TieneArmaduraCazador(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo TieneArmaduraCazador_Err
    
        

100     If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        
102         If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Subtipo = 3 Then ' Aguante hardcodear números :D
104             TieneArmaduraCazador = True
            End If
        
        End If

        
        Exit Function

TieneArmaduraCazador_Err:
106     Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.TieneArmaduraCazador", Erl)

        
End Function

Public Sub SetModoConsulta(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/06/10
        '
        '***************************************************

        Dim sndNick As String

100     With UserList(UserIndex)
102         sndNick = .name
    
104         If .flags.EnConsulta Then
106             sndNick = sndNick & " [CONSULTA]"
            
            Else

108             If .GuildIndex > 0 Then
110                 sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
                End If

            End If

112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .Faccion.Status, sndNick))

        End With

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As eHeading)

    On Error GoTo Handler

    With UserList(UserIndex)

        ' Elegimos un lado al azar
        Dim R As Integer
        R = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

        ' Roto el heading original hacia ese lado
        Heading = RotateHeading(Heading, R)

        ' Intento moverlo para ese lado
        If MoveUserChar(UserIndex, Heading) Then
            ' Le aviso al usuario que fue movido
            Call WriteForceCharMove(UserIndex, Heading)
            Exit Sub
        End If
        
        ' Si falló, intento moverlo para el lado opuesto
        Heading = InvertHeading(Heading)
        If MoveUserChar(UserIndex, Heading) Then
            ' Le aviso al usuario que fue movido
            Call WriteForceCharMove(UserIndex, Heading)
            Exit Sub
        End If
        
        ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
        Dim NuevaPos As WorldPos
        Call ClosestLegalPos(.pos, NuevaPos, .flags.Navegando, .flags.Navegando = 0)
        Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)

    End With

    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.MoveUserToSide", Erl)
    Resume Next
End Sub

' Autor: WyroX - 02/03/2021
' Quita parálisis, veneno, invisibilidad, estupidez, mimetismo, deja de descansar, de meditar y de ocultarse; y quita otros estados obsoletos (por si acaso)
Public Sub LimpiarEstadosAlterados(ByVal UserIndex As Integer)

    On Error GoTo Handler
    
    With UserList(UserIndex)

        '<<<< Envenenamiento >>>>
        .flags.Envenenado = 0
        
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
        
        '<<<< Invisible >>>>
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
        End If
        
        '<<<< Mimetismo >>>>
        If .flags.Mimetizado > 0 Then
        
            If .flags.Navegando Then
            
                If .flags.Muerto = 0 Then
                    .Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
                Else
                    .Char.Body = iFragataFantasmal
                End If

                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
                
            Else
            
                .Char.Body = .CharMimetizado.Body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                
            End If
            
            .Counters.Mimetismo = 0
            .flags.Mimetizado = e_EstadoMimetismo.Desactivado
        End If
        
        '<<<< Estados obsoletos >>>>
        .flags.Ahogandose = 0
        .flags.Incinerado = 0
        
    End With
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.LimpiarEstadosAlterados", Erl)
    Resume Next

End Sub

Public Sub DevolverPosAnterior(ByVal UserIndex As Integer)

    With UserList(UserIndex).flags
        Call WarpToLegalPos(UserIndex, .LastPos.Map, .LastPos.x, .LastPos.y, True)
    End With

End Sub

Public Function ActualizarVelocidadDeUsuario(ByVal UserIndex As Integer) As Single
    On Error GoTo ActualizarVelocidadDeUsuario_Err
    
    Dim velocidad As Single, modificadorItem As Single, modificadorHechizo As Single
    
    velocidad = VelocidadNormal
    modificadorItem = 1
    modificadorHechizo = 1
    
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then
            velocidad = VelocidadMuerto
            GoTo UpdateSpeed ' Los muertos no tienen modificadores de velocidad
        End If
        
        ' El traje para nadar es considerado barco, de subtipo = 0
        If (.flags.Navegando + .flags.Nadando > 0) And (.Invent.BarcoObjIndex > 0) Then
            modificadorItem = ObjData(.Invent.BarcoObjIndex).velocidad
        End If
        
        If (.flags.Montado = 1) And (.Invent.MonturaObjIndex > 0) Then
            modificadorItem = ObjData(.Invent.MonturaObjIndex).velocidad
        End If
        
        ' Algun hechizo le afecto la velocidad
        If .flags.VelocidadHechizada > 0 Then
            modificadorHechizo = .flags.VelocidadHechizada
        End If
        
        velocidad = VelocidadNormal * modificadorItem * modificadorHechizo
        
UpdateSpeed:
        .Char.speeding = velocidad
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        Call WriteVelocidadToggle(UserIndex)
     
    End With

    Exit Function
    
ActualizarVelocidadDeUsuario_Err:
    Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.CalcularVelocidad_Err", Erl)
    Resume Next
End Function

