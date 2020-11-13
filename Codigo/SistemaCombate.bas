Attribute VB_Name = "SistemaCombate"
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Dise�o y correcci�n del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat

Option Explicit

Public Const MAXDISTANCIAARCO  As Byte = 18

Public Const MAXDISTANCIAMAGIA As Byte = 18

Function ModificadorEvasion(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorEvasion_Err
        

100     ModificadorEvasion = ModClase(clase).Evasion

        
        Exit Function

ModificadorEvasion_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModificadorEvasion", Erl)
        Resume Next
        
End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueArmas_Err
        

100     ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

        
        Exit Function

ModificadorPoderAtaqueArmas_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModificadorPoderAtaqueArmas", Erl)
        Resume Next
        
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueProyectiles_Err
        
    
100     ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

        
        Exit Function

ModificadorPoderAtaqueProyectiles_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModificadorPoderAtaqueProyectiles", Erl)
        Resume Next
        
End Function

Function ModicadorDa�oClaseArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDa�oClaseArmas_Err
        
    
100     ModicadorDa�oClaseArmas = ModClase(clase).Da�oArmas

        
        Exit Function

ModicadorDa�oClaseArmas_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorDa�oClaseArmas", Erl)
        Resume Next
        
End Function

Function ModicadorDa�oClaseWrestling(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDa�oClaseWrestling_Err
        
        
100     ModicadorDa�oClaseWrestling = ModClase(clase).Da�oWrestling

        
        Exit Function

ModicadorDa�oClaseWrestling_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorDa�oClaseWrestling", Erl)
        Resume Next
        
End Function

Function ModicadorDa�oClaseProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDa�oClaseProyectiles_Err
        
        
100     ModicadorDa�oClaseProyectiles = ModClase(clase).Da�oProyectiles

        
        Exit Function

ModicadorDa�oClaseProyectiles_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorDa�oClaseProyectiles", Erl)
        Resume Next
        
End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModEvasionDeEscudoClase_Err
        

100     ModEvasionDeEscudoClase = ModClase(clase).Escudo

        
        Exit Function

ModEvasionDeEscudoClase_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModEvasionDeEscudoClase", Erl)
        Resume Next
        
End Function

Function Minimo(ByVal a As Single, ByVal b As Single) As Single
        
        On Error GoTo Minimo_Err
        

100     If a > b Then
102         Minimo = b
            Else:
104         Minimo = a

        End If

        
        Exit Function

Minimo_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.Minimo", Erl)
        Resume Next
        
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
        
        On Error GoTo MinimoInt_Err
        

100     If a > b Then
102         MinimoInt = b
            Else:
104         MinimoInt = a

        End If

        
        Exit Function

MinimoInt_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.MinimoInt", Erl)
        Resume Next
        
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
        
        On Error GoTo Maximo_Err
        

100     If a > b Then
102         Maximo = a
            Else:
104         Maximo = b

        End If

        
        Exit Function

Maximo_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.Maximo", Erl)
        Resume Next
        
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
        
        On Error GoTo MaximoInt_Err
        

100     If a > b Then
102         MaximoInt = a
            Else:
104         MaximoInt = b

        End If

        
        Exit Function

MaximoInt_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.MaximoInt", Erl)
        Resume Next
        
End Function

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasionEscudo_Err
        

100     PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

        
        Exit Function

PoderEvasionEscudo_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderEvasionEscudo", Erl)
        Resume Next
        
End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasion_Err
        

        Dim lTemp As Long

100     With UserList(UserIndex)
102         lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorEvasion(.clase)
       
104         PoderEvasion = (lTemp + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))

        End With

        
        Exit Function

PoderEvasion_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderEvasion", Erl)
        Resume Next
        
End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderAtaqueArma_Err
        

        Dim PoderAtaqueTemp As Long

100     If UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 31 Then
102         PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
104     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 61 Then
106         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
108     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 91 Then
110         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
        Else
112         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))

        End If

114     PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

        
        Exit Function

PoderAtaqueArma_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderAtaqueArma", Erl)
        Resume Next
        
End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderAtaqueProyectil_Err
        

        Dim PoderAtaqueTemp As Long

100     If UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
102         PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
104     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
106         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
108     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
110         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
        Else
112         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))

        End If

114     PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

        
        Exit Function

PoderAtaqueProyectil_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderAtaqueProyectil", Erl)
        Resume Next
        
End Function

Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderAtaqueWrestling_Err
        

        Dim PoderAtaqueTemp As Long

100     If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 31 Then
102         PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
104     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 61 Then
106         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
108     ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 91 Then
110         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
        Else
112         PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))

        End If

114     PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

        
        Exit Function

PoderAtaqueWrestling_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderAtaqueWrestling", Erl)
        Resume Next
        
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo UserImpactoNpc_Err
        

        Dim PoderAtaque As Long

        Dim Arma        As Integer

        Dim proyectil   As Boolean

        Dim ProbExito   As Long

100     Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex

102     If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

104     If Arma > 0 Then 'Usando un arma
106         If proyectil Then
108             PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Else
110             PoderAtaque = PoderAtaqueArma(UserIndex)

            End If

        Else 'Peleando con pu�os
112         PoderAtaque = PoderAtaqueWrestling(UserIndex)

        End If

114     ProbExito = Maximo(10, Minimo(90, 70 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.1)))

116     UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

118     If UserImpactoNpc Then
120         If Arma <> 0 Then
122             If proyectil Then
124                 Call SubirSkill(UserIndex, Proyectiles)
                Else
126                 Call SubirSkill(UserIndex, Armas)

                End If

            Else
128             Call SubirSkill(UserIndex, Wrestling)

            End If

        End If

        
        Exit Function

UserImpactoNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserImpactoNpc", Erl)
        Resume Next
        
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo NpcImpacto_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Revisa si un NPC logra impactar a un user o no
        '03/15/2006 Maraxus - Evit� una divisi�n por cero que eliminaba NPCs
        '*************************************************
        Dim Rechazo           As Boolean

        Dim ProbRechazo       As Long

        Dim ProbExito         As Long

        Dim UserEvasion       As Long

        Dim NpcPoderAtaque    As Long

        Dim PoderEvasioEscudo As Long

        Dim SkillTacticas     As Long

        Dim SkillDefensa      As Long

100     UserEvasion = PoderEvasion(UserIndex)
102     NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
104     PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

106     SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
108     SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)

        'Esta usando un escudo ???
110     If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

112     ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.2)))

114     NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

        ' el usuario esta usando un escudo ???
116     If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
118         If Not NpcImpacto Then
120             If SkillDefensa + SkillTacticas > 0 Then  'Evitamos divisi�n por cero
122                 ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
124                 Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

126                 If Rechazo = True Then
                        'Se rechazo el ataque con el escudo
128                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

130                     If UserList(UserIndex).ChatCombate = 1 Then
132                         Call WriteBlockedWithShieldUser(UserIndex)

                        End If

134                     Call SubirSkill(UserIndex, Defensa)

                        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 88, 0))
                    End If

                End If

            End If

        End If

        
        Exit Function

NpcImpacto_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcImpacto", Erl)
        Resume Next
        
End Function

Public Function CalcularDa�o(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
        
        On Error GoTo CalcularDa�o_Err
        

        Dim Da�oArma As Long, Da�oUsuario As Long, Arma As ObjData, ModifClase As Single

        Dim proyectil As ObjData

        Dim Da�oMaxArma As Long

        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        Dim matoDragon As Boolean

100     matoDragon = False

102     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
104         Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
            ' Ataca a un npc?
106         If NpcIndex > 0 Then

                'Usa la mata Dragones?
108             If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
110                 ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
            
112                 If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
114                     Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
116                     Da�oMaxArma = Arma.MaxHit
118                     matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                    Else ' Sino es Dragon da�o es 1
120                     Da�oArma = 1
122                     Da�oMaxArma = 1

                    End If

                Else ' da�o comun

124                 If Arma.proyectil = 1 Then
126                     ModifClase = ModicadorDa�oClaseProyectiles(UserList(UserIndex).clase)
128                     Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
130                     Da�oMaxArma = Arma.MaxHit

132                     If Arma.Municion = 1 Then
134                         proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
136                         Da�oArma = Da�oArma * 1.35
138                         Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)
140                         Da�oMaxArma = Arma.MaxHit
142                         Da�oMaxArma = Da�oMaxArma * 1.35

                        End If

                    Else
144                     ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
146                     Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
148                     Da�oArma = Da�oArma * 1.35
150                     Da�oMaxArma = Arma.MaxHit
152                     Da�oMaxArma = Da�oMaxArma * 1.35

                    End If

                End If
    
            Else ' Ataca usuario

154             If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
156                 ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
158                 Da�oArma = 1 ' Si usa la espada mataDragones da�o es 1
160                 Da�oMaxArma = 1
                Else

162                 If Arma.proyectil = 1 Then
164                     ModifClase = ModicadorDa�oClaseProyectiles(UserList(UserIndex).clase)
166                     Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
168                     Da�oMaxArma = Arma.MaxHit
                
170                     If Arma.Municion = 1 Then
172                         proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
174                         Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)
176                         Da�oMaxArma = Arma.MaxHit

                        End If

                    Else
178                     ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
180                     Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
182                     Da�oMaxArma = Arma.MaxHit

                    End If

                End If

            End If

        Else

            'Pablo (ToxicWaste)
184         If UserList(UserIndex).Invent.NudilloSlot = 0 Then
186             ModifClase = ModicadorDa�oClaseWrestling(UserList(UserIndex).clase)
188             Da�oArma = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHit)
190             Da�oMaxArma = UserList(UserIndex).Stats.MaxHit
            Else
    
192             ModifClase = ModicadorDa�oClaseArmas(UserList(UserIndex).clase)
194             Arma = ObjData(UserList(UserIndex).Invent.NudilloObjIndex)
196             Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
198             Da�oMaxArma = Arma.MaxHit

            End If

        End If

200     If UserList(UserIndex).Invent.MagicoObjIndex = 707 And NpcIndex = 0 Then
202         Da�oUsuario = RandomNumber((UserList(UserIndex).Stats.MinHIT - ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CuantoAumento), (UserList(UserIndex).Stats.MaxHit - ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CuantoAumento))
        Else
204         Da�oUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHit)

        End If

        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
206     If matoDragon Then
208         CalcularDa�o = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
210         CalcularDa�o = ((3 * Da�oArma) + ((Da�oMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(Fuerza) - 15))) + Da�oUsuario) * ModifClase
    
            'CalcularDa�o = ((3 * 14) + ((14 / 5) * 20) + Da�oUsuario) * ModifClase
            'CalcularDa�o = (42 + (56 + 104) * ModifClase
            'CalcularDa�o = 202 * 0.95  = 191      - defensas
    
            'CalcularDa�o = 136
        End If

        
        Exit Function

CalcularDa�o_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDa�o", Erl)
        Resume Next
        
End Function

Public Sub UserDa�oNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        
        On Error GoTo UserDa�oNpc_Err
        

        Dim da�o As Long

        Dim j As Integer

        Dim apuda�o As Integer
    
100     da�o = CalcularDa�o(UserIndex, NpcIndex)
    
        'esta navegando? si es asi le sumamos el da�o del barco
102     If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then da�o = da�o + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)

104     If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then da�o = da�o + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxHit)

106     If PuedeApu�alar(UserIndex) Then
108         Call SubirSkill(UserIndex, Apu�alar)
110         apuda�o = Apu�alarFunction(UserIndex, NpcIndex, 0, da�o)

            da�o = da�o + apuda�o
        End If
    
112     da�o = da�o - Npclist(NpcIndex).Stats.def
    
114     If da�o < 0 Then da�o = 0
    
        '[KEVIN]
    
        'If UserList(UserIndex).ChatCombate = 1 Then
        '    Call WriteUserHitNPC(UserIndex, da�o)
        'End If
    
116     If apuda�o > 0 Then
118         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead("�" & da�o & "!", Npclist(NpcIndex).Char.CharIndex, &HFFFF00))

120         If UserList(UserIndex).ChatCombate = 1 Then
                'Call WriteConsoleMsg(UserIndex, "Has apu�alado la criatura por " & apuda�o, FontTypeNames.FONTTYPE_FIGHT)
            
122             Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, da�o)

            End If

        Else
124         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(da�o, Npclist(NpcIndex).Char.CharIndex))

        End If
    
126     Call CalcularDarExp(UserIndex, NpcIndex, da�o)
128     Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - da�o
        '[/KEVIN]
     
130     If Npclist(NpcIndex).Stats.MinHp <= 0 Then
            
            ' Si era un Dragon perdemos la espada mataDragones
132         If Npclist(NpcIndex).NPCtype = DRAGON Then

                'Si tiene equipada la matadracos se la sacamos
134             If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
136                 Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)

                End If

                ' If Npclist(NpcIndex).Stats.MaxHp > 100000 Then Call LogDesarrollo(UserList(UserIndex).name & " mat� un drag�n")
            End If
        
138         Call MuereNpc(NpcIndex, UserIndex)

        End If

        
        Exit Sub

UserDa�oNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserDa�oNpc", Erl)
        Resume Next
        
End Sub

Public Sub NpcDa�o(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo NpcDa�o_Err
        

        Dim da�o As Integer, Lugar As Integer, absorbido As Integer

        Dim antda�o As Integer, defbarco As Integer

        Dim obj As ObjData
    
100     da�o = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHit)
102     antda�o = da�o
    
104     If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
106         obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
108         defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

        End If
    
        Dim defMontura As Integer

110     If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
112         obj = ObjData(UserList(UserIndex).Invent.MonturaObjIndex)
114         defMontura = RandomNumber(obj.MinDef, obj.MaxDef)

        End If
    
116     Lugar = RandomNumber(1, 6)
    
118     Select Case Lugar

            Case PartesCuerpo.bCabeza

                'Si tiene casco absorbe el golpe
120             If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
122                 obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
124                 absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
126                 absorbido = absorbido + defbarco
128                 da�o = da�o - absorbido

130                 If da�o < 1 Then da�o = 1

                End If

132         Case Else

                'Si tiene armadura absorbe el golpe
134             If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then

                    Dim Obj2 As ObjData

136                 obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)

138                 If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
140                     Obj2 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
142                     absorbido = RandomNumber(obj.MinDef + Obj2.MinDef, obj.MaxDef + Obj2.MaxDef)
                    Else
144                     absorbido = RandomNumber(obj.MinDef, obj.MaxDef)

                    End If

146                 absorbido = absorbido + defbarco
148                 da�o = da�o - absorbido

150                 If da�o < 1 Then da�o = 1

                End If

        End Select
    
152     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectOverHead(da�o, UserList(UserIndex).Char.CharIndex))

154     If UserList(UserIndex).ChatCombate = 1 Then
156         Call WriteNPCHitUser(UserIndex, Lugar, da�o)

        End If

158     If UserList(UserIndex).flags.Privilegios And PlayerType.user Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - da�o
    
160     If UserList(UserIndex).flags.Meditando Then
162         If da�o > Fix(UserList(UserIndex).Stats.MinHp / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
164             UserList(UserIndex).flags.Meditando = False
166             Call WriteLocaleMsg(UserIndex, "123", FontTypeNames.FONTTYPE_INFO)
168             UserList(UserIndex).Char.FX = 0
170             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
            End If

        End If
    
        'Muere el usuario
172     If UserList(UserIndex).Stats.MinHp <= 0 Then
    
174         Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
        
            'Si lo mato un guardia
176         If Status(UserIndex) = 2 And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then

                ' Call RestarCriminalidad(UserIndex)
178             If Status(UserIndex) < 2 And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)

            End If
        
180         Call UserDie(UserIndex)
    
        End If

        
        Exit Sub

NpcDa�o_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcDa�o", Erl)
        Resume Next
        
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo NpcAtacaUser_Err
        

100     If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
102     If (Not UserList(UserIndex).flags.Privilegios And PlayerType.user) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
        ' El npc puede atacar ???
    
104     If IntervaloPermiteAtacarNPC(NpcIndex) Then
    
106         NpcAtacaUser = True

108         If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    
110         If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
        Else
112         NpcAtacaUser = False
            Exit Function

        End If
    
114     Npclist(NpcIndex).CanAttack = 0
    
116     If Npclist(NpcIndex).flags.Snd1 > 0 Then
118         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))

        End If
    
120     If NpcImpacto(NpcIndex, UserIndex) Then
122         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        
124         If UserList(UserIndex).flags.Navegando = 0 Or UserList(UserIndex).flags.Montado = 0 Then
126             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))

            End If
        
128         Call NpcDa�o(NpcIndex, UserIndex)
130         Call WriteUpdateHP(UserIndex)

            '�Puede envenenar?
132         If Npclist(NpcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, Npclist(NpcIndex).Veneno)
        Else
134         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharSwing(Npclist(NpcIndex).Char.CharIndex, False))

        End If

        '-----Tal vez suba los skills------
136     Call SubirSkill(UserIndex, Tacticas)
    
        'Controla el nivel del usuario
138     Call CheckUserLevel(UserIndex)

        
        Exit Function

NpcAtacaUser_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcAtacaUser", Erl)
        Resume Next
        
End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
        
        On Error GoTo NpcImpactoNpc_Err
        

        Dim PoderAtt  As Long, PoderEva As Long

        Dim ProbExito As Long

100     PoderAtt = Npclist(Atacante).PoderAtaque
102     PoderEva = Npclist(Victima).PoderEvasion
104     ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtt - PoderEva) * 0.4)))
106     NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

        
        Exit Function

NpcImpactoNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcImpactoNpc", Erl)
        Resume Next
        
End Function

Public Sub NpcDa�oNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
        
        On Error GoTo NpcDa�oNpc_Err
        

        Dim da�o As Integer

        Dim ANpc As npc

100     ANpc = Npclist(Atacante)
    
102     da�o = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHit)
104     Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - da�o
    
106     If Npclist(Victima).Stats.MinHp < 1 Then
        
108         If LenB(Npclist(Atacante).flags.AttackedBy) <> 0 Then
110             Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
112             Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
            Else
114             Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement

            End If
        
        End If

        
        Exit Sub

NpcDa�oNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcDa�oNpc", Erl)
        Resume Next
        
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
        
        On Error GoTo NpcAtacaNpc_Err
        
 
        ' El npc puede atacar ???
100     If IntervaloPermiteAtacarNPC(Atacante) Then
102         Npclist(Atacante).CanAttack = 0

104         If cambiarMOvimiento Then
106             Npclist(Victima).TargetNPC = Atacante
108             Npclist(Victima).Movement = TipoAI.NpcAtacaNpc

            End If

        Else
            Exit Sub

        End If

110     If Npclist(Atacante).flags.Snd1 > 0 Then
112         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1, Npclist(Atacante).Pos.x, Npclist(Atacante).Pos.Y))

        End If

114     If NpcImpactoNpc(Atacante, Victima) Then
    
116         If Npclist(Victima).flags.Snd2 > 0 Then
118             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
            Else
120             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))

            End If

122         Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
    
124         Call NpcDa�oNpc(Atacante, Victima)
    
        Else
126         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageCharSwing(Npclist(Atacante).Char.CharIndex, False, True))

        End If

        
        Exit Sub

NpcAtacaNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcAtacaNpc", Erl)
        Resume Next
        
End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        
        On Error GoTo UsuarioAtacaNpc_Err
        

100     If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            Exit Sub

        End If
    
102     If UserList(UserIndex).flags.invisible = 0 Then
104         Call NPCAtacado(NpcIndex, UserIndex)

        End If

106     If UserImpactoNpc(UserIndex, NpcIndex) Then
        
108         If Npclist(NpcIndex).flags.Snd2 > 0 Then
110             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
            Else
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))

            End If

114         If UserList(UserIndex).flags.Paraliza = 1 And Npclist(NpcIndex).flags.Paralizado = 0 Then

                Dim Probabilidad As Byte
    
116             Probabilidad = RandomNumber(1, 4)

118             If Probabilidad = 1 Then
120                 If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
122                     Npclist(NpcIndex).flags.Paralizado = 1
                        
124                     Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado / 3

126                     If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
128                         Call WriteLocaleMsg(UserIndex, "136", FontTypeNames.FONTTYPE_FIGHT)

                        End If

130                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.CharIndex, 8, 0))
                                     
                    Else

132                     If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
134                         Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If

            End If

136         Call UserDa�oNpc(UserIndex, NpcIndex)
       
        Else
138         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))

        End If

        
        Exit Sub

UsuarioAtacaNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacaNpc", Erl)
        Resume Next
        
End Sub

Public Function UsuarioAtacaNpcFunction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Byte
        
        On Error GoTo UsuarioAtacaNpcFunction_Err
        

100     If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
102         UsuarioAtacaNpcFunction = 0
            Exit Function

        End If
    
104     Call NPCAtacado(NpcIndex, UserIndex)
    
106     If UserImpactoNpc(UserIndex, NpcIndex) Then
108         Call UserDa�oNpc(UserIndex, NpcIndex)
110         UsuarioAtacaNpcFunction = 1
        Else
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))
114         UsuarioAtacaNpcFunction = 2

        End If

        
        Exit Function

UsuarioAtacaNpcFunction_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacaNpcFunction", Erl)
        Resume Next
        
End Function

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
        
        On Error GoTo UsuarioAtaca_Err
        

        'Check bow's interval
100     If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
        'Check Spell-Attack interval
102     If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

        'Check Attack interval
104     If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub

        'Quitamos stamina
106     If UserList(UserIndex).Stats.MinSta < 10 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
108         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
110     Call QuitarSta(UserIndex, RandomNumber(1, 10))
    
112     If UserList(UserIndex).Counters.Trabajando Then
114         Call WriteMacroTrabajoToggle(UserIndex, False)

        End If
        
116     If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

        'Movimiento de arma
118     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))
     
        'UserList(UserIndex).flags.PuedeAtacar = 0
    
        Dim AttackPos As WorldPos

120     AttackPos = UserList(UserIndex).Pos
122     Call HeadtoPos(UserList(UserIndex).Char.heading, AttackPos)
       
        'Exit if not legal
124     If AttackPos.x >= XMinMapSize And AttackPos.x <= XMaxMapSize And AttackPos.Y >= YMinMapSize And AttackPos.Y <= YMaxMapSize Then

            Dim Index As Integer

126         Index = MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).UserIndex
            
            'Look for user
128         If Index > 0 Then
130             Call UsuarioAtacaUsuario(UserIndex, Index)
132             Call WriteUpdateUserStats(UserIndex)
134             Call WriteUpdateUserStats(Index)
                Exit Sub

            End If
            
            'Look for NPC
136         If MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex > 0 Then
            
138             If Npclist(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex).Attackable Then
                    
140                 Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex)
142                 Call WriteUpdateUserStats(UserIndex)
                Else
            
144                 Call WriteConsoleMsg(UserIndex, "No pod�s atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If

        Else
146         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))

        End If

        
        Exit Sub

UsuarioAtaca_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtaca", Erl)
        Resume Next
        
End Sub

Public Function UsuarioImpacto(ByVal atacanteindex As Integer, ByVal victimaindex As Integer) As Boolean
        
        On Error GoTo UsuarioImpacto_Err
        
    
        Dim ProbRechazo            As Long

        Dim Rechazo                As Boolean

        Dim ProbExito              As Long

        Dim PoderAtaque            As Long

        Dim UserPoderEvasion       As Long

        Dim UserPoderEvasionEscudo As Long

        Dim Arma                   As Integer

        Dim proyectil              As Boolean

        Dim SkillTacticas          As Long

        Dim SkillDefensa           As Long
    
100     If UserList(atacanteindex).flags.GolpeCertero = 1 Then
102         UsuarioImpacto = True
104         UserList(atacanteindex).flags.GolpeCertero = 0
            Exit Function

        End If
    
106     SkillTacticas = UserList(victimaindex).Stats.UserSkills(eSkill.Tacticas)
108     SkillDefensa = UserList(victimaindex).Stats.UserSkills(eSkill.Defensa)
    
110     Arma = UserList(atacanteindex).Invent.WeaponEqpObjIndex

112     If Arma > 0 Then
114         proyectil = ObjData(Arma).proyectil = 1
        Else
116         proyectil = False

        End If
    
        'Calculamos el poder de evasion...
118     UserPoderEvasion = PoderEvasion(victimaindex)
    
120     If UserList(victimaindex).Invent.EscudoEqpObjIndex > 0 Then
122         UserPoderEvasionEscudo = PoderEvasionEscudo(victimaindex)
124         UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
        Else
126         UserPoderEvasionEscudo = 0

        End If
    
        'Esta usando un arma ???
128     If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
        
130         If proyectil Then
132             PoderAtaque = PoderAtaqueProyectil(atacanteindex)
            Else
134             PoderAtaque = PoderAtaqueArma(atacanteindex)

            End If

136         ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))
       
        Else
138         PoderAtaque = PoderAtaqueWrestling(atacanteindex)
140         ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))
        
        End If

142     UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
        ' el usuario esta usando un escudo ???
144     If UserList(victimaindex).Invent.EscudoEqpObjIndex > 0 Then
        
            'Fallo ???
146         If UsuarioImpacto = False Then
148             ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
150             Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
          
152             If Rechazo = True Then
                    'Se rechazo el ataque con el escudo
154                 Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessagePlayWave(SND_ESCUDO, UserList(victimaindex).Pos.x, UserList(victimaindex).Pos.Y))
156                 Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEscudoMov(UserList(victimaindex).Char.CharIndex))

158                 If UserList(atacanteindex).ChatCombate = 1 Then
160                     Call WriteBlockedWithShieldOther(atacanteindex)

                    End If

162                 If UserList(victimaindex).ChatCombate = 1 Then
164                     Call WriteBlockedWithShieldUser(victimaindex)

                    End If

166                 Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 88, 0))
                
168                 Call SubirSkill(victimaindex, Defensa)

                End If

            End If

        End If
        
170     If UsuarioImpacto Then
172         If Arma > 0 Then
174             If Not proyectil Then
176                 Call SubirSkill(atacanteindex, Armas)
                Else
178                 Call SubirSkill(atacanteindex, Proyectiles)

                End If

            Else
180             Call SubirSkill(atacanteindex, Wrestling)

            End If

        End If

        
        Exit Function

UsuarioImpacto_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioImpacto", Erl)
        Resume Next
        
End Function

Public Sub UsuarioAtacaUsuario(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UsuarioAtacaUsuario_Err
        

        Dim Probabilidad As Byte

        Dim HuboEfecto   As Boolean
    
100     If Not PuedeAtacar(atacanteindex, victimaindex) Then Exit Sub
    
102     If Distancia(UserList(atacanteindex).Pos, UserList(victimaindex).Pos) > MAXDISTANCIAARCO Then
104         Call WriteLocaleMsg(atacanteindex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(atacanteindex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

106     HuboEfecto = False
    
108     Call UsuarioAtacadoPorUsuario(atacanteindex, victimaindex)
    
110     If UsuarioImpacto(atacanteindex, victimaindex) Then
112         Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessagePlayWave(SND_IMPACTO, UserList(atacanteindex).Pos.x, UserList(atacanteindex).Pos.Y))
        
114         If UserList(victimaindex).flags.Navegando = 0 Or UserList(victimaindex).flags.Montado = 0 Then
116             Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, FXSANGRE, 0))

            End If
        
118         If UserList(atacanteindex).flags.incinera = 1 Then
120             Probabilidad = RandomNumber(1, 6)

122             If Probabilidad = 1 Then
124                 If UserList(victimaindex).flags.Incinerado = 0 Then
126                     UserList(victimaindex).flags.Incinerado = 1

128                     If UserList(victimaindex).ChatCombate = 1 Then
130                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

132                     If UserList(atacanteindex).ChatCombate = 1 Then
134                         Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

136                     HuboEfecto = True
138                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Incinerar, 100, False))

                    End If

                End If

            End If
    
140         If UserList(atacanteindex).flags.Envenena > 0 And Not HuboEfecto Then
142             Probabilidad = RandomNumber(1, 2)
    
144             If Probabilidad = 1 Then
146                 If UserList(victimaindex).flags.Envenenado = 0 Then
148                     UserList(victimaindex).flags.Envenenado = UserList(atacanteindex).flags.Envenena

150                     If UserList(victimaindex).ChatCombate = 1 Then
152                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                    
154                     If UserList(atacanteindex).ChatCombate = 1 Then
156                         Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

158                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Envenena, 100, False))

                    End If

                End If

            End If
        
160         If UserList(atacanteindex).flags.Paraliza = 1 And Not HuboEfecto Then
162             Probabilidad = RandomNumber(1, 5)

164             If Probabilidad = 1 Then
166                 If UserList(victimaindex).flags.Paralizado = 0 Then
168                     UserList(victimaindex).flags.Paralizado = 1
170                     UserList(victimaindex).Counters.Paralisis = 6
172                     Call WriteParalizeOK(victimaindex)
                        Rem   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
174                     If UserList(victimaindex).ChatCombate = 1 Then
176                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha paralizado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                    
178                     If UserList(atacanteindex).ChatCombate = 1 Then
180                         Call WriteConsoleMsg(atacanteindex, "Has paralizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

                        'Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Paralizar, 100, False))
182                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 8, 0))

                    End If

                End If

            End If
        
184         Call UserDa�oUser(atacanteindex, victimaindex)

        Else
186         Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessageCharSwing(UserList(atacanteindex).Char.CharIndex))

        End If

        
        Exit Sub

UsuarioAtacaUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacaUsuario", Erl)
        Resume Next
        
End Sub

Public Sub UserDa�oUser(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UserDa�oUser_Err
        

        Dim da�o As Long, antda�o As Integer

        Dim Lugar    As Integer, absorbido As Long

        Dim defbarco As Integer

        Dim apuda�o As Integer

        Dim obj As ObjData
    
100     da�o = CalcularDa�o(atacanteindex)
102     antda�o = da�o

104     If PuedeApu�alar(atacanteindex) Then
106         apuda�o = Apu�alarFunction(atacanteindex, 0, victimaindex, da�o)
108         da�o = da�o + apuda�o
110         antda�o = da�o

        End If

112     Call UserDa�oEspecial(atacanteindex, victimaindex)
    
114     If UserList(atacanteindex).flags.Navegando = 1 And UserList(atacanteindex).Invent.BarcoObjIndex > 0 Then
116         obj = ObjData(UserList(atacanteindex).Invent.BarcoObjIndex)
118         da�o = da�o + RandomNumber(obj.MinHIT, obj.MaxHit)

        End If
    
120     If UserList(victimaindex).flags.Navegando = 1 And UserList(victimaindex).Invent.BarcoObjIndex > 0 Then
122         obj = ObjData(UserList(victimaindex).Invent.BarcoObjIndex)
124         defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

        End If
    
126     If UserList(atacanteindex).flags.Montado = 1 And UserList(atacanteindex).Invent.MonturaObjIndex > 0 Then
128         obj = ObjData(UserList(atacanteindex).Invent.MonturaObjIndex)
130         da�o = da�o + RandomNumber(obj.MinHIT, obj.MaxHit)

        End If
    
132     If UserList(victimaindex).flags.Montado = 1 And UserList(victimaindex).Invent.MonturaObjIndex > 0 Then
134         obj = ObjData(UserList(victimaindex).Invent.MonturaObjIndex)
136         defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

        End If
    
        Dim Resist As Byte

138     If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
140         Resist = ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).Refuerzo

        End If
    
142     Lugar = RandomNumber(1, 6)
    
144     Select Case Lugar
      
            Case PartesCuerpo.bCabeza

                'Si tiene casco absorbe el golpe
146             If UserList(victimaindex).Invent.CascoEqpObjIndex > 0 Then
148                 obj = ObjData(UserList(victimaindex).Invent.CascoEqpObjIndex)
150                 absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
152                 absorbido = absorbido + defbarco - Resist
154                 da�o = da�o - absorbido

156                 If da�o < 0 Then da�o = 1

                End If

158         Case Else

                'Si tiene armadura absorbe el golpe
160             If UserList(victimaindex).Invent.ArmourEqpObjIndex > 0 Then
162                 obj = ObjData(UserList(victimaindex).Invent.ArmourEqpObjIndex)

                    Dim Obj2 As ObjData

164                 If UserList(victimaindex).Invent.EscudoEqpObjIndex Then
166                     Obj2 = ObjData(UserList(victimaindex).Invent.EscudoEqpObjIndex)
168                     absorbido = RandomNumber(obj.MinDef + Obj2.MinDef, obj.MaxDef + Obj2.MaxDef)
                    Else
170                     absorbido = RandomNumber(obj.MinDef, obj.MaxDef)

                    End If

172                 absorbido = absorbido + defbarco - Resist
174                 da�o = da�o - absorbido

176                 If da�o < 0 Then da�o = 1

                End If

        End Select
    
178     If apuda�o > 0 Then
180         Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEfectOverHead("�" & da�o & "!", UserList(victimaindex).Char.CharIndex, &HFFFF00))
            
            If UserList(atacanteindex).ChatCombate = 1 Then
                Call WriteConsoleMsg(atacanteindex, "Has apu�alado a " & UserList(victimaindex).name & " por " & da�o & ".", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            If UserList(victimaindex).ChatCombate = 1 Then
                Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha apu�alado por " & da�o & ".", FontTypeNames.FONTTYPE_FIGHT)
            End If

            Call WriteEfectToScreen(victimaindex, &H3C3CFF, 200, True)
            Call WriteEfectToScreen(atacanteindex, &H3C3CFF, 150, True)
            
            Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 89, 0))
        Else
182         Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEfectOverHead(da�o, UserList(victimaindex).Char.CharIndex))
        End If
    
184     If UserList(atacanteindex).ChatCombate = 1 Then
186         Call WriteUserHittedUser(atacanteindex, Lugar, UserList(victimaindex).Char.CharIndex, da�o - apuda�o)

        End If
    
188     If UserList(victimaindex).ChatCombate = 1 Then
190         Call WriteUserHittedByUser(victimaindex, Lugar, UserList(atacanteindex).Char.CharIndex, da�o - apuda�o)

        End If

192     UserList(victimaindex).Stats.MinHp = UserList(victimaindex).Stats.MinHp - da�o
    
194     If UserList(atacanteindex).flags.Hambre = 0 And UserList(atacanteindex).flags.Sed = 0 Then

            'Si usa un arma quizas suba "Combate con armas"
196         If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
198             If ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
200                 Call SubirSkill(atacanteindex, Proyectiles)
                Else
                    'Sube combate con armas.
202                 Call SubirSkill(atacanteindex, Armas)

                End If

            Else
                'sino tal vez lucha libre
204             Call SubirSkill(atacanteindex, Wrestling)

            End If
            
206         Call SubirSkill(victimaindex, Tacticas)

208         If PuedeApu�alar(atacanteindex) Then
210             Call SubirSkill(atacanteindex, Apu�alar)

            End If
    
            'e intenta dar un golpe cr�tico [Pablo (ToxicWaste)]
            ' Call DoGolpeCritico(atacanteindex, 0, victimaindex, da�o)
        End If
    
212     If UserList(victimaindex).Stats.MinHp <= 0 Then
    
            'Store it!
214         Call Statistics.StoreFrag(atacanteindex, victimaindex)
        
216         Call ContarMuerte(victimaindex, atacanteindex)
    
218         Call ActStats(victimaindex, atacanteindex)
        Else
            'Est� vivo - Actualizamos el HP
    
220         Call WriteUpdateHP(victimaindex)

        End If
    
        'Controla el nivel del usuario
222     Call CheckUserLevel(atacanteindex)
    
    

        
        Exit Sub

UserDa�oUser_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserDa�oUser", Erl)
        Resume Next
        
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Autor: Unknown
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************
        
        On Error GoTo UsuarioAtacadoPorUsuario_Err
        

100     If UserList(VictimIndex).flags.Meditando Then
102         UserList(VictimIndex).flags.Meditando = False
104         Call WriteLocaleMsg(VictimIndex, "123", FontTypeNames.FONTTYPE_INFO)
106         UserList(VictimIndex).Char.FX = 0
108         Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.CharIndex, 0))
        End If
    
110     If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
        Dim EraCriminal As Byte
    
112     UserList(VictimIndex).Counters.TiempoDeMapeo = 3
114     UserList(attackerIndex).Counters.TiempoDeMapeo = 3
    
116     If Status(attackerIndex) = 1 And Status(VictimIndex) = 1 Or Status(VictimIndex) = 3 Then
118         Call VolverCriminal(attackerIndex)

        End If
    
120     EraCriminal = Status(attackerIndex)
    
122     If EraCriminal = 2 And Status(attackerIndex) < 2 Then
124         Call RefreshCharStatus(attackerIndex)
126     ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
128         Call RefreshCharStatus(attackerIndex)

        End If

130     If Status(attackerIndex) = 2 Then If UserList(attackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(attackerIndex)
    
        'If UserList(VictimIndex).Familiar.Existe = 1 Then
        '  If UserList(VictimIndex).Familiar.Invocado = 1 Then
        '  Npclist(UserList(VictimIndex).Familiar.Id).flags.AttackedBy = UserList(attackerIndex).name
        '  Npclist(UserList(VictimIndex).Familiar.Id).Movement = TipoAI.NPCDEFENSA
        '  Npclist(UserList(VictimIndex).Familiar.Id).Hostile = 1
        ' End If
        ' End If
    
        'Si la victima esta saliendo se cancela la salida
132     Call CancelExit(VictimIndex)
    

        
        Exit Sub

UsuarioAtacadoPorUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacadoPorUsuario", Erl)
        Resume Next
        
End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
        
        On Error GoTo PuedeAtacar_Err
        

        '***************************************************
        'Autor: Unknown
        'Last Modification: 24/01/2007
        'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
        '24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
        '***************************************************
        Dim T    As eTrigger6

        Dim rank As Integer

        'MUY importante el orden de estos "IF"...

        'Estas muerto no podes atacar
100     If UserList(attackerIndex).flags.Muerto = 1 Then
102         Call WriteLocaleMsg(attackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "No pod�s atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacar = False
            Exit Function

        End If

        'No podes atacar a alguien muerto
106     If UserList(VictimIndex).flags.Muerto = 1 Then
108         Call WriteConsoleMsg(attackerIndex, "No pod�s atacar a un espiritu.", FontTypeNames.FONTTYPE_INFOIAO)
110         PuedeAtacar = False
            Exit Function

        End If

112     If UserList(attackerIndex).flags.Maldicion = 1 Then
114         Call WriteConsoleMsg(attackerIndex, "�Estas maldito! No podes atacar.", FontTypeNames.FONTTYPE_INFOIAO)
116         PuedeAtacar = False
            Exit Function

        End If

        'Es miembro del grupo?
118     If UserList(attackerIndex).Grupo.EnGrupo = True Then

            Dim i As Byte

120         For i = 1 To UserList(UserList(attackerIndex).Grupo.Lider).Grupo.CantidadMiembros
    
122             If UserList(UserList(attackerIndex).Grupo.Lider).Grupo.Miembros(i) = VictimIndex Then
124                 PuedeAtacar = False
126                 Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Function

                End If

128         Next i

        End If

        'Estamos en una Arena? o un trigger zona segura?
130     T = TriggerZonaPelea(attackerIndex, VictimIndex)

132     If T = eTrigger6.TRIGGER6_PERMITE Then
134         PuedeAtacar = True
            Exit Function
136     ElseIf T = eTrigger6.TRIGGER6_PROHIBE Then
138         PuedeAtacar = False
            Exit Function
140     ElseIf T = eTrigger6.TRIGGER6_AUSENTE Then

            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            ' If Not UserList(VictimIndex).flags.Privilegios And PlayerType.User Then
            '   If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
            ' PuedeAtacar = False
            '    Exit Function
            ' End If
        End If

        'Consejeros no pueden atacar
        'If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
        '    PuedeAtacar = False
        '    Exit Sub
        'End If

142     If UserList(attackerIndex).GuildIndex <> 0 Then
144         If UserList(attackerIndex).flags.SeguroClan Then
146             If UserList(attackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
148                 Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu clan. Para hacerlo debes desactivar el seguro de clan.", FontTypeNames.FONTTYPE_INFOIAO)
150                 PuedeAtacar = False
                    Exit Function

                End If

            End If

        End If

        'Estas queriendo atacar a un GM?
152     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

154     If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
156         If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
158         PuedeAtacar = False
            Exit Function

        End If

        'Sos un Armada atacando un ciudadano?
160     If (Status(VictimIndex) = 1) And (esArmada(attackerIndex)) Then
162         Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
164         PuedeAtacar = False
            Exit Function

        End If

        'Tenes puesto el seguro?
166     If UserList(attackerIndex).flags.Seguro Then
168         If Status(VictimIndex) = 1 Then
170             Call WriteConsoleMsg(attackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
172             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Es un ciuda queriando atacar un imperial?

174     If UserList(attackerIndex).flags.Seguro Then
176         If (Status(attackerIndex) = 1) And (esArmada(VictimIndex)) Then
178             Call WriteConsoleMsg(attackerIndex, "Los ciudadanos no pueden atacar a los soldados imperiales.", FontTypeNames.FONTTYPE_WARNING)
180             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Estas en un Mapa Seguro?
182     If MapInfo(UserList(VictimIndex).Pos.Map).Seguro = 1 Then
184         If esArmada(attackerIndex) Then
186             If UserList(attackerIndex).Faccion.RecompensasReal > 11 Then
188                 If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
190                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podr�s defenderte.", FontTypeNames.FONTTYPE_WARNING)
192                     PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

194         If esCaos(attackerIndex) Then
196             If UserList(attackerIndex).Faccion.RecompensasCaos > 11 Then
198                 If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
200                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podr�s defenderte.", FontTypeNames.FONTTYPE_WARNING)
202                     PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

204         Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
206         PuedeAtacar = False
            Exit Function

        End If

        'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
208     If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.x, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(attackerIndex).Pos.Map, UserList(attackerIndex).Pos.x, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
210         Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
212         PuedeAtacar = False
            Exit Function

        End If

214     PuedeAtacar = True

        
        Exit Function

PuedeAtacar_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PuedeAtacar", Erl)
        Resume Next
        
End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
        '***************************************************
        'Autor: Unknown Author (Original version)
        'Returns True if AttackerIndex can attack the NpcIndex
        'Last Modification: 24/01/2007
        '24/01/2007 Pablo (ToxicWaste) - Orden y correcci�n de ataque sobre una mascota y guardias
        '14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
        'esta funci�n para todo lo referente a ataque a un NPC. Ya sea Magia, F�sico o a Distancia.
        '***************************************************
        
        On Error GoTo PuedeAtacarNPC_Err
        

        'Estas muerto?
100     If UserList(attackerIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(attackerIndex, "No pod�s atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
102         Call WriteLocaleMsg(attackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacarNPC = False
            Exit Function

        End If

        'Sos consejero?
106     If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
            'No pueden atacar NPC los Consejeros.
108         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es una criatura atacable?
110     If Npclist(NpcIndex).Attackable = 0 Then
            'No es una criatura atacable
112         Call WriteConsoleMsg(attackerIndex, "No pod�s atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
114         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es valida la distancia a la cual estamos atacando?
116     If Distancia(UserList(attackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
118         Call WriteLocaleMsg(attackerIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
120         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es una criatura No-Hostil?
122     If Npclist(NpcIndex).Hostile = 0 Then
            'Es una criatura No-Hostil.
            'Es Guardia del Caos?

124         If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then

                'Lo quiere atacar un caos?
126             If esCaos(attackerIndex) Then
128                 Call WriteConsoleMsg(attackerIndex, "No pod�s atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
130                 PuedeAtacarNPC = False
                    Exit Function

                End If

132             If Status(attackerIndex) = 1 Then
134                 PuedeAtacarNPC = True
                    Exit Function

                End If
        
            End If

            'Es guardia Real?
136         If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                'Lo quiere atacar un Armada?
        
138             If esCaos(attackerIndex) Then
140                 PuedeAtacarNPC = True
                    Exit Function

                End If
        
142             If esArmada(attackerIndex) Then
144                 Call WriteConsoleMsg(attackerIndex, "No pod�s atacar Guardias Reales siendo Armada Real", FontTypeNames.FONTTYPE_INFO)
146                 PuedeAtacarNPC = False
                    Exit Function

                End If
        
                'Tienes el seguro puesto?
148             If UserList(attackerIndex).flags.Seguro And Status(attackerIndex) = 1 Then
150                 Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para poder Atacar Guardias Reales utilizando /seg", FontTypeNames.FONTTYPE_INFO)
152                 PuedeAtacarNPC = False
                    Exit Function
                Else
154                 Call WriteConsoleMsg(attackerIndex, "Atacaste un Guardia Real! Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
156                 Call VolverCriminal(attackerIndex)
158                 PuedeAtacarNPC = True
                    Exit Function

                End If

            End If

        End If

160     PuedeAtacarNPC = True

        
        Exit Function

PuedeAtacarNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PuedeAtacarNPC", Erl)
        Resume Next
        
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDa�o As Long)
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        
        On Error GoTo CalcularDarExp_Err
        

100     If UserList(UserIndex).Grupo.EnGrupo Then
102         Call CalcularDarExpGrupal(UserIndex, NpcIndex, ElDa�o)
        Else

            Dim ExpaDar As Long
    
            '[Nacho] Chekeamos que las variables sean validas para las operaciones
104         If ElDa�o <= 0 Then ElDa�o = 0
106         If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
108         If ElDa�o > Npclist(NpcIndex).Stats.MinHp Then ElDa�o = Npclist(NpcIndex).Stats.MinHp
    
            '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
110         ExpaDar = CLng((ElDa�o) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

112         If ExpaDar <= 0 Then Exit Sub

            '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
114         If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
116             ExpaDar = Npclist(NpcIndex).flags.ExpCount
118             Npclist(NpcIndex).flags.ExpCount = 0
            Else
120             Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar

            End If
    
122         If ExpMult > 0 Then
124             ExpaDar = ExpaDar * ExpMult * UserList(UserIndex).flags.ScrollExp
    
            End If
    
126         If UserList(UserIndex).donador.activo = 1 Then
128             ExpaDar = ExpaDar * 1.1

            End If
    
            '[Nacho] Le damos la exp al user
130         If ExpaDar > 0 Then
132             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
134                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

136                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

138                 Call WriteUpdateExp(UserIndex)
140                 Call CheckUserLevel(UserIndex)

                End If
            
142             Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, ExpaDar, 6)

            End If

        End If

        
        Exit Sub

CalcularDarExp_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDarExp", Erl)
        Resume Next
        
End Sub

Sub CalcularDarExpGrupal(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDa�o As Long)
        
        On Error GoTo CalcularDarExpGrupal_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim ExpaDar           As Long

        Dim BonificacionGrupo As Single

        'If UserList(UserIndex).Grupo.EnGrupo Then
        '[Nacho] Chekeamos que las variables sean validas para las operaciones
100     If ElDa�o <= 0 Then ElDa�o = 0
102     If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
104     If ElDa�o > Npclist(NpcIndex).Stats.MinHp Then ElDa�o = Npclist(NpcIndex).Stats.MinHp
    
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
106     ExpaDar = CLng((ElDa�o) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

108     If ExpaDar <= 0 Then Exit Sub

        '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
110     If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
112         ExpaDar = Npclist(NpcIndex).flags.ExpCount
114         Npclist(NpcIndex).flags.ExpCount = 0
        Else
116         Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar

        End If
    
118     Select Case UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
            Case 1
120             BonificacionGrupo = 1

122         Case 2
124             BonificacionGrupo = 1.2

126         Case 3
128             BonificacionGrupo = 1.4

130         Case 4
132             BonificacionGrupo = 1.6

134         Case 5
136             BonificacionGrupo = 1.8

138         Case 6
140             BonificacionGrupo = 2
                
        End Select
 
142     If ExpMult > 0 Then
144         ExpaDar = ExpaDar * ExpMult
        
        End If
    
        Dim expbackup As Long

146     expbackup = ExpaDar
148     ExpaDar = ExpaDar * BonificacionGrupo

        Dim i     As Byte

        Dim Index As Byte

150     expbackup = expbackup / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
152     ExpaDar = ExpaDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
        Dim ExpUser As Long
    
154     For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
156         Index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)

158         If UserList(Index).flags.Muerto = 0 Then
160             If UserList(UserIndex).Pos.Map = UserList(Index).Pos.Map Then
162                 If ExpaDar > 0 Then
164                     ExpUser = 0

166                     If UserList(Index).donador.activo = 1 Then
168                         ExpUser = ExpaDar * 1.1
                        Else
170                         ExpUser = ExpaDar

                        End If
                    
172                     ExpUser = ExpUser * UserList(Index).flags.ScrollExp
                
174                     If UserList(Index).Stats.ELV < STAT_MAXELV Then
176                         UserList(Index).Stats.Exp = UserList(Index).Stats.Exp + ExpUser

178                         If UserList(Index).Stats.Exp > MAXEXP Then UserList(Index).Stats.Exp = MAXEXP

180                         If UserList(Index).ChatCombate = 1 Then
182                             Call WriteLocaleMsg(Index, "141", FontTypeNames.FONTTYPE_EXP, ExpUser)

                            End If

184                         Call WriteUpdateExp(Index)
186                         Call CheckUserLevel(Index)

                        End If

                    End If

                Else

                    'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
188                 If UserList(Index).ChatCombate = 1 Then
190                     Call WriteLocaleMsg(Index, "69", FontTypeNames.FONTTYPE_New_GRUPO)

                    End If

192                 If expbackup > 0 Then
194                     If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
196                         UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + expbackup

198                         If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

200                         If UserList(UserIndex).ChatCombate = 1 Then
202                             Call WriteConsoleMsg(UserIndex, UserList(Index).name & " estas demasiado lejos de tu grupo, has ganado " & expbackup & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)

                            End If

204                         Call CheckUserLevel(UserIndex)
206                         Call WriteUpdateExp(UserIndex)

                        End If

                    End If

                End If

            Else

208             If UserList(Index).ChatCombate = 1 Then
210                 Call WriteConsoleMsg(Index, "Estas muerto, no has ganado experencia del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

                End If

212             If expbackup > 0 Then
214                 If UserList(Index).Stats.ELV < STAT_MAXELV Then
216                     UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + expbackup

218                     If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

220                     If UserList(UserIndex).ChatCombate = 1 Then
222                         Call WriteConsoleMsg(UserIndex, UserList(Index).name & " estas muerto, has ganado " & expbackup & " puntos de experiencia correspondientes a el.", FontTypeNames.FONTTYPE_EXP)

                        End If

224                     Call CheckUserLevel(UserIndex)
226                     Call WriteUpdateExp(UserIndex)

                    End If

                End If

            End If

228     Next i

        'Else
        '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, experencia perdida.", FontTypeNames.FONTTYPE_New_GRUPO)
        'End If

        
        Exit Sub

CalcularDarExpGrupal_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDarExpGrupal", Erl)
        Resume Next
        
End Sub

Sub CalcularDarOroGrupal(ByVal UserIndex As Integer, ByVal GiveGold As Long)
        
        On Error GoTo CalcularDarOroGrupal_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim OroDar            As Long

        Dim BonificacionGrupo As Single

        'If UserList(UserIndex).Grupo.EnGrupo Then
    
100     Select Case UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
            Case 1
102             BonificacionGrupo = 1

104         Case 2
106             BonificacionGrupo = 1.2

108         Case 3
110             BonificacionGrupo = 1.4

112         Case 4
114             BonificacionGrupo = 1.6

116         Case 5
118             BonificacionGrupo = 1.8

120         Case 6
122             BonificacionGrupo = 2
                
        End Select
 
124     OroDar = GiveGold * OroMult
    
126     If OroDar > 0 Then
128         OroDar = OroDar * BonificacionGrupo
        
        End If
    
        Dim orobackup As Long
    
130     orobackup = OroDar

        Dim i     As Byte

        Dim Index As Byte

132     OroDar = OroDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

134     For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
136         Index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)

138         If UserList(Index).flags.Muerto = 0 Then
140             If UserList(UserIndex).Pos.Map = UserList(Index).Pos.Map Then
142                 If OroDar > 0 Then
                    
144                     OroDar = orobackup * UserList(Index).flags.ScrollOro
                
146                     UserList(Index).Stats.GLD = UserList(Index).Stats.GLD + OroDar
                        
148                     If UserList(Index).ChatCombate = 1 Then
150                         Call WriteConsoleMsg(Index, "�El grupo ha ganado " & OroDar & " monedas de oro!", FontTypeNames.FONTTYPE_New_GRUPO)

                        End If
                        
152                     Call WriteUpdateGold(Index)
                        
                    End If

                Else

                    'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
                    'Call WriteLocaleMsg(Index, "69", FontTypeNames.FONTTYPE_INFOIAO)
                End If

            Else

                '  Call WriteConsoleMsg(Index, "Estas muerto, no has ganado oro del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
            End If

154     Next i

        'Else
        '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, oro perdido.", FontTypeNames.FONTTYPE_New_GRUPO)
        'End If

        
        Exit Sub

CalcularDarOroGrupal_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDarOroGrupal", Erl)
        Resume Next
        
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

    'TODO: Pero que rebuscado!!
    'Nigo:  Te lo redise�e, pero no te borro el TODO para que lo revises.
    On Error GoTo Errhandler

    Dim tOrg As eTrigger

    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.x, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.x, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE

        End If

    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE

    End If

    Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)

End Function

Sub UserIncinera(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UserIncinera_Err
        

        Dim ArmaObjInd As Integer, ObjInd As Integer

        Dim num        As Long
 
100     ArmaObjInd = UserList(atacanteindex).Invent.WeaponEqpObjIndex
102     ObjInd = 0
 
104     If ArmaObjInd > 0 Then
106         If ObjData(ArmaObjInd).proyectil = 0 Then
108             ObjInd = ArmaObjInd
            Else
110             ObjInd = UserList(atacanteindex).Invent.MunicionEqpObjIndex

            End If
   
112         If ObjInd > 0 Then
114             If (ObjData(ObjInd).incinera = 1) Then
116                 num = RandomNumber(1, 6)
           
118                 If num < 6 Then
120                     UserList(victimaindex).flags.Incinerado = 1
122                     Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)
124                     Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                End If

            End If

        End If
 
    

        
        Exit Sub

UserIncinera_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserIncinera", Erl)
        Resume Next
        
End Sub

Sub UserDa�oEspecial(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UserDa�oEspecial_Err
        

        Dim ArmaObjInd As Integer, ObjInd As Integer

        Dim HuboEfecto As Boolean

        Dim num        As Long

100     HuboEfecto = False
102     ArmaObjInd = UserList(atacanteindex).Invent.WeaponEqpObjIndex
104     ObjInd = 0

106     If ArmaObjInd = 0 Then
108         ArmaObjInd = UserList(atacanteindex).Invent.NudilloObjIndex

        End If

110     If ArmaObjInd > 0 Then
112         If ObjData(ArmaObjInd).proyectil = 0 Then
114             ObjInd = ArmaObjInd
            Else
116             ObjInd = UserList(atacanteindex).Invent.MunicionEqpObjIndex

            End If
    
118         If ObjInd > 0 Then
120             If (ObjData(ObjInd).Envenena > 0) And Not HuboEfecto Then
122                 num = RandomNumber(1, 100)
            
124                 If num < 30 Then
126                     UserList(victimaindex).flags.Envenenado = ObjData(ObjInd).Envenena
128                     Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
130                     Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)
132                     HuboEfecto = True

                    End If

                End If
        
134             If (ObjData(ObjInd).incinera > 0) And Not HuboEfecto Then
136                 num = RandomNumber(1, 100)
            
138                 If num < 10 Then
140                     UserList(victimaindex).flags.Incinerado = 1
142                     Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)
144                     Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)
146                     HuboEfecto = True

                    End If

                End If
        
148             If (ObjData(ObjInd).Paraliza > 0) And Not HuboEfecto Then
150                 num = RandomNumber(1, 100)

152                 If num < 10 Then
154                     If UserList(victimaindex).flags.Paralizado = 0 Then
156                         UserList(victimaindex).flags.Paralizado = 1
158                         UserList(victimaindex).Counters.Paralisis = 6
160                         Call WriteParalizeOK(victimaindex)
162                         Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 8, 0))
                    
164                         If UserList(victimaindex).ChatCombate = 1 Then
166                             Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha paralizado!!", FontTypeNames.FONTTYPE_FIGHT)

                            End If
                    
168                         If UserList(atacanteindex).ChatCombate = 1 Then
170                             Call WriteConsoleMsg(atacanteindex, "Has paralizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                            End If

172                         HuboEfecto = True
                    
                        End If

                    End If

                End If
        
174             If (ObjData(ObjInd).Estupidiza > 0) And Not HuboEfecto Then
176                 num = RandomNumber(1, 100)

178                 If num < 8 Then
180                     If UserList(victimaindex).flags.Estupidez = 0 Then
182                         UserList(victimaindex).flags.Estupidez = 1
184                         UserList(victimaindex).Counters.Estupidez = 5

                        End If
                
186                     If UserList(victimaindex).ChatCombate = 1 Then
188                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha estupidizado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

190                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, 30, 30, False))
                
192                     If UserList(atacanteindex).ChatCombate = 1 Then
194                         Call WriteConsoleMsg(atacanteindex, "Has estupidizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

196                     Call WriteDumb(victimaindex)

                    End If

                End If

            End If

        End If

    

        
        Exit Sub

UserDa�oEspecial_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserDa�oEspecial", Erl)
        Resume Next
        
End Sub

