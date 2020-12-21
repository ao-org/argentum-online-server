Attribute VB_Name = "SistemaCombate"
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
'
'Diseño y corrección del modulo de combate por
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
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModificadorEvasion", Erl)
104     Resume Next
        
End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueArmas_Err
        

100     ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

        
        Exit Function

ModificadorPoderAtaqueArmas_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModificadorPoderAtaqueArmas", Erl)
104     Resume Next
        
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueProyectiles_Err
        
    
100     ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

        
        Exit Function

ModificadorPoderAtaqueProyectiles_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModificadorPoderAtaqueProyectiles", Erl)
104     Resume Next
        
End Function

Function ModicadorDañoClaseArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseArmas_Err
        
    
100     ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

        
        Exit Function

ModicadorDañoClaseArmas_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorDañoClaseArmas", Erl)
104     Resume Next
        
End Function
Function ModicadorApuñalarClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorApuñalarClase_Err
        
    
100     ModicadorApuñalarClase = ModClase(clase).ModApuñalar

        
        Exit Function

ModicadorApuñalarClase_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorApuñalarClase", Erl)
104     Resume Next
        
End Function

Function ModicadorDañoClaseWrestling(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseWrestling_Err
        
        
100     ModicadorDañoClaseWrestling = ModClase(clase).DañoWrestling

        
        Exit Function

ModicadorDañoClaseWrestling_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorDañoClaseWrestling", Erl)
104     Resume Next
        
End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseProyectiles_Err
        
        
100     ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

        
        Exit Function

ModicadorDañoClaseProyectiles_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModicadorDañoClaseProyectiles", Erl)
104     Resume Next
        
End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModEvasionDeEscudoClase_Err
        

100     ModEvasionDeEscudoClase = ModClase(clase).Escudo

        
        Exit Function

ModEvasionDeEscudoClase_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.ModEvasionDeEscudoClase", Erl)
104     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.Minimo", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.MinimoInt", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.Maximo", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.MaximoInt", Erl)
108     Resume Next
        
End Function

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasionEscudo_Err
        

100     PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

        
        Exit Function

PoderEvasionEscudo_Err:
102     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderEvasionEscudo", Erl)
104     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderEvasion", Erl)
108     Resume Next
        
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
116     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderAtaqueArma", Erl)
118     Resume Next
        
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
116     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderAtaqueProyectil", Erl)
118     Resume Next
        
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
116     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PoderAtaqueWrestling", Erl)
118     Resume Next
        
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

        Else 'Peleando con puños
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
130     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserImpactoNpc", Erl)
132     Resume Next
        
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo NpcImpacto_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Revisa si un NPC logra impactar a un user o no
        '03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
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
120             If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
122                 ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
124                 Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

126                 If Rechazo = True Then
                        'Se rechazo el ataque con el escudo
128                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

130                     If UserList(UserIndex).ChatCombate = 1 Then
132                         Call WriteBlockedWithShieldUser(UserIndex)

                        End If

                        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 88, 0))
                    End If

                End If

            End If
            
134         Call SubirSkill(UserIndex, Defensa)

        End If

        
        Exit Function

NpcImpacto_Err:
136     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcImpacto", Erl)
138     Resume Next
        
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
        
        On Error GoTo CalcularDaño_Err
        

        Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single

        Dim proyectil As ObjData

        Dim DañoMaxArma As Long

        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        Dim matoDragon As Boolean

100     matoDragon = False

102     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
104         Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
            ' Ataca a un npc?
106         If NpcIndex > 0 Then

                'Usa la mata Dragones?
108             If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
110                 ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            
112                 If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
114                     DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
116                     DañoMaxArma = Arma.MaxHit
118                     matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                    Else ' Sino es Dragon daño es 1
120                     DañoArma = 1
122                     DañoMaxArma = 1

                    End If

                Else ' daño comun

124                 If Arma.proyectil = 1 Then
126                     ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
128                     DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
130                     DañoMaxArma = Arma.MaxHit

132                     If Arma.Municion = 1 Then
134                         proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
136                         DañoArma = DañoArma
138                         DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)
140                         DañoMaxArma = Arma.MaxHit
142                         DañoMaxArma = DañoMaxArma

                        End If

                    Else
144                     ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
146                     DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
148                     DañoArma = DañoArma
150                     DañoMaxArma = Arma.MaxHit
152                     DañoMaxArma = DañoMaxArma

                    End If

                End If
    
            Else ' Ataca usuario

154             If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
156                 ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
158                 DañoArma = 1 ' Si usa la espada mataDragones daño es 1
160                 DañoMaxArma = 1
                Else

162                 If Arma.proyectil = 1 Then
164                     ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
166                     DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
168                     DañoMaxArma = Arma.MaxHit
                
170                     If Arma.Municion = 1 Then
172                         proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
174                         DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)
176                         DañoMaxArma = Arma.MaxHit

                        End If

                    Else
178                     ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
180                     DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
182                     DañoMaxArma = Arma.MaxHit

                    End If

                End If

            End If

        Else

            'Pablo (ToxicWaste)
184         If UserList(UserIndex).Invent.NudilloSlot = 0 Then
186             ModifClase = ModicadorDañoClaseWrestling(UserList(UserIndex).clase)
188             DañoArma = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHit)
190             DañoMaxArma = UserList(UserIndex).Stats.MaxHit
            Else
    
192             ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
194             Arma = ObjData(UserList(UserIndex).Invent.NudilloObjIndex)
196             DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
198             DañoMaxArma = Arma.MaxHit

            End If

        End If

200     If UserList(UserIndex).Invent.MagicoObjIndex = 707 And NpcIndex = 0 Then
202         DañoUsuario = RandomNumber((UserList(UserIndex).Stats.MinHIT - ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CuantoAumento), (UserList(UserIndex).Stats.MaxHit - ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CuantoAumento))
        Else
204         DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHit)

        End If

        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
206     If matoDragon Then
208         CalcularDaño = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
210         CalcularDaño = ((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(Fuerza) - 15))) + DañoUsuario) * ModifClase
    
            'CalcularDaño = ((3 * 14) + ((14 / 5) * 20) + DañoUsuario) * ModifClase
            'CalcularDaño = (42 + (56 + 104) * ModifClase
            'CalcularDaño = 202 * 0.95  = 191      - defensas
    
            'CalcularDaño = 136
        End If

        
        Exit Function

CalcularDaño_Err:
212     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDaño", Erl)
214     Resume Next
        
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        
        On Error GoTo UserDañoNpc_Err
        

        Dim daño As Long

        Dim j As Integer

        Dim apudaño As Integer
    
100     daño = CalcularDaño(UserIndex, NpcIndex)
    
        'esta navegando? si es asi le sumamos el daño del barco
102     If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)

104     If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxHit)

106     If PuedeApuñalar(UserIndex) Then
108         Call SubirSkill(UserIndex, Apuñalar)
110         apudaño = ApuñalarFunction(UserIndex, NpcIndex, 0, daño)

112         daño = daño + apudaño
        End If
    
114     daño = daño - Npclist(NpcIndex).Stats.def
    
116     If daño < 0 Then daño = 0
    
        '[KEVIN]
    
        'If UserList(UserIndex).ChatCombate = 1 Then
        '    Call WriteUserHitNPC(UserIndex, daño)
        'End If
    
118     If apudaño > 0 Then
120         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead("¡" & daño & "!", Npclist(NpcIndex).Char.CharIndex, vbYellow))

122         If UserList(UserIndex).ChatCombate = 1 Then
                'Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & apudaño, FontTypeNames.FONTTYPE_FIGHT)
            
124             Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, daño)

            End If

        Else
126         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex))

        End If
        
128     If UserList(UserIndex).ChatCombate = 1 Then
130         Call WriteConsoleMsg(UserIndex, "Le has causado " & daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
132     Call CalcularDarExp(UserIndex, NpcIndex, daño)
134     Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
        '[/KEVIN]
     
136     If Npclist(NpcIndex).Stats.MinHp <= 0 Then
            
            ' Si era un Dragon perdemos la espada mataDragones
138         If Npclist(NpcIndex).NPCtype = DRAGON Then

                'Si tiene equipada la matadracos se la sacamos
140             If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
142                 Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)

                End If

                ' If Npclist(NpcIndex).Stats.MaxHp > 100000 Then Call LogDesarrollo(UserList(UserIndex).name & " mató un dragón")
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
144         For j = 1 To MAXMASCOTAS
146             If UserList(UserIndex).MascotasIndex(j) > 0 Then
148                 If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
150                     Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
152                     Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                    End If
                End If
154         Next j
        
156         Call MuereNpc(NpcIndex, UserIndex)

        End If

        
        Exit Sub

UserDañoNpc_Err:
158     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserDañoNpc", Erl)
160     Resume Next
        
End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo NpcDaño_Err
        

        Dim daño As Integer, Lugar As Integer, absorbido As Integer

        Dim antdaño As Integer, defbarco As Integer

        Dim obj As ObjData
    
100     daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHit)
102     antdaño = daño
    
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
128                 daño = daño - absorbido

130                 If daño < 1 Then daño = 1

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
148                 daño = daño - absorbido

150                 If daño < 1 Then daño = 1

                End If

        End Select
    
152     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectOverHead(daño, UserList(UserIndex).Char.CharIndex))

154     If UserList(UserIndex).ChatCombate = 1 Then
156         Call WriteNPCHitUser(UserIndex, Lugar, daño)

        End If

158     If UserList(UserIndex).flags.Privilegios And PlayerType.user Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - daño
    
160     If UserList(UserIndex).flags.Meditando Then
162         If daño > Fix(UserList(UserIndex).Stats.MinHp / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
164             UserList(UserIndex).flags.Meditando = False
166             UserList(UserIndex).Char.FX = 0
168             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
            End If

        End If
    
        'Muere el usuario
170     If UserList(UserIndex).Stats.MinHp <= 0 Then
    
172         Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
        
            'Si lo mato un guardia
174         If Status(UserIndex) = 2 And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then

                ' Call RestarCriminalidad(UserIndex)
176             If Status(UserIndex) < 2 And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)

            End If
            
178         If Npclist(NpcIndex).MaestroUser > 0 Then
180             Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
182             If Npclist(NpcIndex).Stats.Alineacion = 0 Then
184                 Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
186                 Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
188                 Npclist(NpcIndex).flags.AttackedBy = vbNullString
                End If
            End If
        
190         Call UserDie(UserIndex)
    
        End If

        
        Exit Sub

NpcDaño_Err:
192     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcDaño", Erl)
194     Resume Next
        
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Heading As eHeading) As Boolean
        
        On Error GoTo NpcAtacaUser_Err
        

100     If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
102     If (Not UserList(UserIndex).flags.Privilegios And PlayerType.user) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
        ' El npc puede atacar ???
    
104     If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
106         NpcAtacaUser = False
            Exit Function
        End If
        
108     If ((MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked And 2 ^ (Heading - 1)) <> 0) Then
110         NpcAtacaUser = False
            Exit Function
        End If

112     NpcAtacaUser = True

114     Call CheckPets(NpcIndex, UserIndex, False)

116     If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    
118     If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex

120     Npclist(NpcIndex).CanAttack = 0
    
122     If Npclist(NpcIndex).flags.Snd1 > 0 Then
124         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        End If
    
126     If NpcImpacto(NpcIndex, UserIndex) Then
128         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
130         If UserList(UserIndex).flags.Navegando = 0 Or UserList(UserIndex).flags.Montado = 0 Then
132             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))

            End If
        
134         Call NpcDaño(NpcIndex, UserIndex)
136         Call WriteUpdateHP(UserIndex)

            '¿Puede envenenar?
138         If Npclist(NpcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, Npclist(NpcIndex).Veneno)
        Else
140         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharSwing(Npclist(NpcIndex).Char.CharIndex, False))

        End If

        '-----Tal vez suba los skills------
142     Call SubirSkill(UserIndex, Tacticas)
    
        'Controla el nivel del usuario
144     Call CheckUserLevel(UserIndex)

        
        Exit Function

NpcAtacaUser_Err:
146     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcAtacaUser", Erl)
148     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcImpactoNpc", Erl)
110     Resume Next
        
End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
        
            On Error GoTo NpcDañoNpc_Err

            Dim daño As Integer
    
100         With Npclist(Atacante)
102             daño = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
104             Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - daño
            
106             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessageEfectOverHead(daño, Npclist(Victima).Char.CharIndex))
            
                ' Mascotas dan experiencia al amo
108             If .MaestroUser > 0 Then
110                 Call CalcularDarExp(.MaestroUser, Victima, daño)
                End If
            
112             If Npclist(Victima).Stats.MinHp < 1 Then
114                 .Movement = .flags.OldMovement
                
116                 If LenB(.flags.AttackedBy) <> 0 Then
118                     .Hostile = .flags.OldHostil
                    End If
                
120                 If .MaestroUser > 0 Then
122                     Call FollowAmo(Atacante)
                    End If
                
124                 Call MuereNpc(Victima, .MaestroUser)
                End If
            End With

        
            Exit Sub

NpcDañoNpc_Err:
126         Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcDañoNpc")
128         Resume Next
        
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
112         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1, Npclist(Atacante).Pos.X, Npclist(Atacante).Pos.Y))

        End If

114     If NpcImpactoNpc(Atacante, Victima) Then
    
116         If Npclist(Victima).flags.Snd2 > 0 Then
118             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
120             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))

            End If

122         Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
    
124         Call NpcDañoNpc(Atacante, Victima)
    
        Else
126         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageCharSwing(Npclist(Atacante).Char.CharIndex, False, True))

        End If

        
        Exit Sub

NpcAtacaNpc_Err:
128     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.NpcAtacaNpc", Erl)
130     Resume Next
        
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
110             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
            Else
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))

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

136         Call UserDañoNpc(UserIndex, NpcIndex)
       
        Else
138         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))

        End If

        
        Exit Sub

UsuarioAtacaNpc_Err:
140     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacaNpc", Erl)
142     Resume Next
        
End Sub

Public Function UsuarioAtacaNpcFunction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Byte
        
        On Error GoTo UsuarioAtacaNpcFunction_Err
        

100     If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
102         UsuarioAtacaNpcFunction = 0
            Exit Function

        End If
    
104     Call NPCAtacado(NpcIndex, UserIndex)
    
106     If UserImpactoNpc(UserIndex, NpcIndex) Then
108         Call UserDañoNpc(UserIndex, NpcIndex)
110         UsuarioAtacaNpcFunction = 1
        Else
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))
114         UsuarioAtacaNpcFunction = 2

        End If

        
        Exit Function

UsuarioAtacaNpcFunction_Err:
116     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacaNpcFunction", Erl)
118     Resume Next
        
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
        
        'Movimiento de arma, solo lo envio si no es GM invisible.
118     If UserList(UserIndex).flags.AdminInvisible = 0 Then
120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))
        End If
        
        'UserList(UserIndex).flags.PuedeAtacar = 0
    
        Dim AttackPos As WorldPos
122         AttackPos = UserList(UserIndex).Pos

124     Call HeadtoPos(UserList(UserIndex).Char.Heading, AttackPos)
       
        'Exit if not legal
126     If AttackPos.X >= XMinMapSize And AttackPos.X <= XMaxMapSize And AttackPos.Y >= YMinMapSize And AttackPos.Y <= YMaxMapSize Then

128         If ((MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).Blocked And 2 ^ (UserList(UserIndex).Char.Heading - 1)) <> 0) Then
130             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))
                Exit Sub
            End If

            Dim index As Integer

132         index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
            
            'Look for user
134         If index > 0 Then
136             Call UsuarioAtacaUsuario(UserIndex, index)
138             Call WriteUpdateUserStats(UserIndex)
140             Call WriteUpdateUserStats(index)
                Exit Sub
            
            'Look for NPC
142         ElseIf MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then

144             index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex
            
146             If Npclist(index).Attackable Then
148                 If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.Map).Seguro = 1 Then
150                     Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                    
152                 Call UsuarioAtacaNpc(UserIndex, index)
154                 Call WriteUpdateUserStats(UserIndex)
                Else
            
156                 Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub
                
            Else
158             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))
            End If

        Else
160         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))
        End If

        
        Exit Sub

UsuarioAtaca_Err:
162     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtaca", Erl)
164     Resume Next
        
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
154                 Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessagePlayWave(SND_ESCUDO, UserList(victimaindex).Pos.X, UserList(victimaindex).Pos.Y))
156                 Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEscudoMov(UserList(victimaindex).Char.CharIndex))

158                 If UserList(atacanteindex).ChatCombate = 1 Then
160                     Call WriteBlockedWithShieldOther(atacanteindex)

                    End If

162                 If UserList(victimaindex).ChatCombate = 1 Then
164                     Call WriteBlockedWithShieldUser(victimaindex)

                    End If

166                 Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 88, 0))

                End If

            End If
            
168         Call SubirSkill(victimaindex, Defensa)

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
182     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioImpacto", Erl)
184     Resume Next
        
End Function

Public Sub UsuarioAtacaUsuario(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UsuarioAtacaUsuario_Err

        Dim Probabilidad As Byte
        Dim HuboEfecto   As Boolean
    
100     If Not PuedeAtacar(atacanteindex, victimaindex) Then Exit Sub
    
102     If Distancia(UserList(atacanteindex).Pos, UserList(victimaindex).Pos) > MAXDISTANCIAARCO Then
104         Call WriteLocaleMsg(atacanteindex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(atacanteindex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

106     HuboEfecto = False
    
108     Call UsuarioAtacadoPorUsuario(atacanteindex, victimaindex)
    
110     If UsuarioImpacto(atacanteindex, victimaindex) Then
112         Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessagePlayWave(SND_IMPACTO, UserList(atacanteindex).Pos.X, UserList(atacanteindex).Pos.Y))
        
114         If UserList(victimaindex).flags.Navegando = 0 Or UserList(victimaindex).flags.Montado = 0 Then
116             Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, FXSANGRE, 0))

            End If
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en accion
118         If UserList(atacanteindex).clase = eClass.Bandit Then
120             Call DoDesequipar(atacanteindex, victimaindex)
                
                'y ahora, el ladron puede llegar a paralizar con el golpe.
122         ElseIf UserList(atacanteindex).clase = eClass.Thief Then
124             Call DoHandInmo(atacanteindex, victimaindex)

            End If
            
126         If UserList(atacanteindex).flags.incinera = 1 Then
128             Probabilidad = RandomNumber(1, 6)

130             If Probabilidad = 1 Then
132                 If UserList(victimaindex).flags.Incinerado = 0 Then
134                     UserList(victimaindex).flags.Incinerado = 1

136                     If UserList(victimaindex).ChatCombate = 1 Then
138                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

140                     If UserList(atacanteindex).ChatCombate = 1 Then
142                         Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

144                     HuboEfecto = True
146                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Incinerar, 100, False))

                    End If

                End If

            End If
    
148         If UserList(atacanteindex).flags.Envenena > 0 And Not HuboEfecto Then
150             Probabilidad = RandomNumber(1, 2)
    
152             If Probabilidad = 1 Then
154                 If UserList(victimaindex).flags.Envenenado = 0 Then
156                     UserList(victimaindex).flags.Envenenado = UserList(atacanteindex).flags.Envenena

158                     If UserList(victimaindex).ChatCombate = 1 Then
160                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                    
162                     If UserList(atacanteindex).ChatCombate = 1 Then
164                         Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

166                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Envenena, 100, False))

                    End If

                End If

            End If
        
168         If UserList(atacanteindex).flags.Paraliza = 1 And Not HuboEfecto Then
170             Probabilidad = RandomNumber(1, 5)

172             If Probabilidad = 1 Then
174                 If UserList(victimaindex).flags.Paralizado = 0 Then
176                     UserList(victimaindex).flags.Paralizado = 1
178                     UserList(victimaindex).Counters.Paralisis = 6
180                     Call WriteParalizeOK(victimaindex)
                        Rem   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
182                     If UserList(victimaindex).ChatCombate = 1 Then
184                         Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha paralizado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                    
186                     If UserList(atacanteindex).ChatCombate = 1 Then
188                         Call WriteConsoleMsg(atacanteindex, "Has paralizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

                        'Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Paralizar, 100, False))
190                     Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 8, 0))

                    End If

                End If

            End If
        
192         Call UserDañoUser(atacanteindex, victimaindex)

        Else
194         Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessageCharSwing(UserList(atacanteindex).Char.CharIndex))

        End If

        
        Exit Sub

UsuarioAtacaUsuario_Err:
196     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacaUsuario", Erl)
198     Resume Next
        
End Sub

Public Sub UserDañoUser(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UserDañoUser_Err
        

        Dim daño As Long, antdaño As Integer

        Dim Lugar    As Integer, absorbido As Long

        Dim defbarco As Integer

        Dim apudaño As Integer

        Dim obj As ObjData
    
100     daño = CalcularDaño(atacanteindex)
102     antdaño = daño

104     If PuedeApuñalar(atacanteindex) Then
106         apudaño = ApuñalarFunction(atacanteindex, 0, victimaindex, daño)
108         daño = daño + apudaño
110         antdaño = daño

        End If

112     Call UserDañoEspecial(atacanteindex, victimaindex)
    
114     If UserList(atacanteindex).flags.Navegando = 1 And UserList(atacanteindex).Invent.BarcoObjIndex > 0 Then
116         obj = ObjData(UserList(atacanteindex).Invent.BarcoObjIndex)
118         daño = daño + RandomNumber(obj.MinHIT, obj.MaxHit)

        End If
    
120     If UserList(victimaindex).flags.Navegando = 1 And UserList(victimaindex).Invent.BarcoObjIndex > 0 Then
122         obj = ObjData(UserList(victimaindex).Invent.BarcoObjIndex)
124         defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

        End If
    
126     If UserList(atacanteindex).flags.Montado = 1 And UserList(atacanteindex).Invent.MonturaObjIndex > 0 Then
128         obj = ObjData(UserList(atacanteindex).Invent.MonturaObjIndex)
130         daño = daño + RandomNumber(obj.MinHIT, obj.MaxHit)

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
154                 daño = daño - absorbido

156                 If daño < 0 Then daño = 1

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
174                 daño = daño - absorbido

176                 If daño < 0 Then daño = 1

                End If

        End Select
    
178     If apudaño > 0 Then
180         Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEfectOverHead("¡" & daño & "!", UserList(victimaindex).Char.CharIndex, vbYellow))
            
182         If UserList(atacanteindex).ChatCombate = 1 Then
184             Call WriteConsoleMsg(atacanteindex, "Has apuñalado a " & UserList(victimaindex).name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
186         If UserList(victimaindex).ChatCombate = 1 Then
188             Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha apuñalado por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
            End If

190         Call WriteEfectToScreen(victimaindex, &H3C3CFF, 200, True)
192         Call WriteEfectToScreen(atacanteindex, &H3C3CFF, 150, True)
            
194         Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 89, 0))
        Else
196         If UserList(atacanteindex).ChatCombate = 1 Then
198             Call WriteUserHittedUser(atacanteindex, Lugar, UserList(victimaindex).Char.CharIndex, daño)
            End If
    
200         If UserList(victimaindex).ChatCombate = 1 Then
202             Call WriteUserHittedByUser(victimaindex, Lugar, UserList(atacanteindex).Char.CharIndex, daño)
            End If
        
204         Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEfectOverHead(daño, UserList(victimaindex).Char.CharIndex))
        End If

206     UserList(victimaindex).Stats.MinHp = UserList(victimaindex).Stats.MinHp - daño
    
208     If UserList(atacanteindex).flags.Hambre = 0 And UserList(atacanteindex).flags.Sed = 0 Then

            'Si usa un arma quizas suba "Combate con armas"
210         If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
212             If ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
214                 Call SubirSkill(atacanteindex, Proyectiles)
                Else
                    'Sube combate con armas.
216                 Call SubirSkill(atacanteindex, Armas)

                End If

            Else
                'sino tal vez lucha libre
218             Call SubirSkill(atacanteindex, Wrestling)

            End If
            
220         Call SubirSkill(victimaindex, Tacticas)

222         If PuedeApuñalar(atacanteindex) Then
224             Call SubirSkill(atacanteindex, Apuñalar)

            End If
    
            'Se intenta dar un golpe crítico [Pablo (ToxicWaste)]
226         Call DoGolpeCritico(atacanteindex, 0, victimaindex, daño)
        End If
    
228     If UserList(victimaindex).Stats.MinHp <= 0 Then
    
            'Store it!
230         Call Statistics.StoreFrag(atacanteindex, victimaindex)
        
232         Call ContarMuerte(victimaindex, atacanteindex)

            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
234         For j = 1 To MAXMASCOTAS
236             If UserList(atacanteindex).MascotasIndex(j) > 0 Then
238                 If Npclist(UserList(atacanteindex).MascotasIndex(j)).Target = victimaindex Then
240                     Npclist(UserList(atacanteindex).MascotasIndex(j)).Target = 0
242                     Call FollowAmo(UserList(atacanteindex).MascotasIndex(j))
                    End If
                End If
244         Next j
    
246         Call ActStats(victimaindex, atacanteindex)
        Else
            'Está vivo - Actualizamos el HP
    
248         Call WriteUpdateHP(victimaindex)

        End If
    
        'Controla el nivel del usuario
250     Call CheckUserLevel(atacanteindex)
    
    

        
        Exit Sub

UserDañoUser_Err:
252     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserDañoUser", Erl)
254     Resume Next
        
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
104         UserList(VictimIndex).Char.FX = 0
106         Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.CharIndex, 0))
        End If
    
108     If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
        Dim EraCriminal As Byte
    
110     UserList(VictimIndex).Counters.TiempoDeMapeo = 3
112     UserList(attackerIndex).Counters.TiempoDeMapeo = 3
    
114     If Status(attackerIndex) = 1 And Status(VictimIndex) = 1 Or Status(VictimIndex) = 3 Then
116         Call VolverCriminal(attackerIndex)

        End If

118     EraCriminal = Status(attackerIndex)
    
120     If EraCriminal = 2 And Status(attackerIndex) < 2 Then
122         Call RefreshCharStatus(attackerIndex)
124     ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
126         Call RefreshCharStatus(attackerIndex)
        End If

128     If Status(attackerIndex) = 2 Then If UserList(attackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(attackerIndex)
    
        'If UserList(VictimIndex).Familiar.Existe = 1 Then
        '  If UserList(VictimIndex).Familiar.Invocado = 1 Then
        '  Npclist(UserList(VictimIndex).Familiar.Id).flags.AttackedBy = UserList(attackerIndex).name
        '  Npclist(UserList(VictimIndex).Familiar.Id).Movement = TipoAI.NPCDEFENSA
        '  Npclist(UserList(VictimIndex).Familiar.Id).Hostile = 1
        ' End If
        ' End If
        
130     Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
132     Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
        'Si la victima esta saliendo se cancela la salida
134     Call CancelExit(VictimIndex)
    

        
        Exit Sub

UsuarioAtacadoPorUsuario_Err:
136     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UsuarioAtacadoPorUsuario", Erl)
138     Resume Next
        
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
            'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacar = False
            Exit Function

        End If

        'No podes atacar a alguien muerto
106     If UserList(VictimIndex).flags.Muerto = 1 Then
108         Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espiritu.", FontTypeNames.FONTTYPE_INFOIAO)
110         PuedeAtacar = False
            Exit Function

        End If

112     If UserList(attackerIndex).flags.Maldicion = 1 Then
114         Call WriteConsoleMsg(attackerIndex, "¡Estas maldito! No podes atacar.", FontTypeNames.FONTTYPE_INFOIAO)
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
        
        'Solo administradores pueden atacar a usuarios (PARA TESTING)
152     If EsGM(attackerIndex) And (UserList(attackerIndex).flags.Privilegios And PlayerType.Admin) = 0 Then
154         PuedeAtacar = False
            Exit Function
        End If
        
        'Estas queriendo atacar a un GM?
156     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

158     If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
160         If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
162         PuedeAtacar = False
            Exit Function

        End If

        'Sos un Armada atacando un ciudadano?
164     If (Status(VictimIndex) = 1) And (esArmada(attackerIndex)) Then
166         Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
168         PuedeAtacar = False
            Exit Function

        End If

        'Tenes puesto el seguro?
170     If UserList(attackerIndex).flags.Seguro Then
172         If Status(VictimIndex) = 1 Then
174             Call WriteConsoleMsg(attackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
176             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Es un ciuda queriando atacar un imperial?
178     If UserList(attackerIndex).flags.Seguro Then
180         If (Status(attackerIndex) = 1) And (esArmada(VictimIndex)) Then
182             Call WriteConsoleMsg(attackerIndex, "Los ciudadanos no pueden atacar a los soldados imperiales.", FontTypeNames.FONTTYPE_WARNING)
184             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Estas en un Mapa Seguro?
186     If MapInfo(UserList(VictimIndex).Pos.Map).Seguro = 1 Then

188         If esArmada(attackerIndex) Then
190             If UserList(attackerIndex).Faccion.RecompensasReal > 11 Then
192                 If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
194                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
196                     PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

198         If esCaos(attackerIndex) Then
200             If UserList(attackerIndex).Faccion.RecompensasCaos > 11 Then
202                 If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
204                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
206                     PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

208         Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
210         PuedeAtacar = False
            Exit Function

        End If

        'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
212     If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(attackerIndex).Pos.Map, UserList(attackerIndex).Pos.X, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
214         Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
216         PuedeAtacar = False
            Exit Function

        End If

218     PuedeAtacar = True

        
        Exit Function

PuedeAtacar_Err:
220     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PuedeAtacar", Erl)
222     Resume Next
        
End Function

Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
        '***************************************************
        'Autor: Unknown Author (Original version)
        'Returns True if AttackerIndex can attack the NpcIndex
        'Last Modification: 24/01/2007
        '24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
        '14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
        'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
        '***************************************************
        
        On Error GoTo PuedeAtacarNPC_Err
        

        'Estas muerto?
100     If UserList(attackerIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
102         Call WriteLocaleMsg(attackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacarNPC = False
            Exit Function

        End If

        'Solo administradores pueden atacar a usuarios (PARA TESTING)
106     If EsGM(attackerIndex) And (UserList(attackerIndex).flags.Privilegios And PlayerType.Admin) = 0 Then
108         PuedeAtacarNPC = False
            Exit Function
        End If
        
        'Es una criatura atacable?
110     If Npclist(NpcIndex).Attackable = 0 Then
            'No es una criatura atacable
112         Call WriteConsoleMsg(attackerIndex, "No podés atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
114         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es valida la distancia a la cual estamos atacando?
116     If Distancia(UserList(attackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
118         Call WriteLocaleMsg(attackerIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
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
128                 Call WriteConsoleMsg(attackerIndex, "No podés atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
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
144                 Call WriteConsoleMsg(attackerIndex, "No podés atacar Guardias Reales siendo Armada Real", FontTypeNames.FONTTYPE_INFO)
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
        
        'Es el NPC mascota de alguien?
160     If Npclist(NpcIndex).MaestroUser > 0 Then
162         If UserList(Npclist(NpcIndex).MaestroUser).Faccion.Status = 1 Then
                'Es mascota de un Ciudadano.
164             If UserList(attackerIndex).Faccion.Status = 1 Then
                    'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
166                 If UserList(attackerIndex).flags.Seguro Then
                        'El atacante tiene el seguro puesto. No puede atacar.
168                     Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro de combate.", FontTypeNames.FONTTYPE_INFO)
170                     PuedeAtacarNPC = False
                        Exit Function
                    Else
                        'El atacante no tiene el seguro puesto. Recibe penalización.
172                     Call WriteConsoleMsg(attackerIndex, "Has atacado la mascota de un ciudadano. Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
174                     Call VolverCriminal(attackerIndex)
176                     PuedeAtacarNPC = True
                        Exit Function
                    End If
                Else
                    'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
178                 If UserList(attackerIndex).flags.Seguro Then
180                     Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro de combate.", FontTypeNames.FONTTYPE_INFO)
182                     PuedeAtacarNPC = False
                        Exit Function
                    End If
                End If
            Else
                'Es mascota de un Criminal.
184             If esCaos(Npclist(NpcIndex).MaestroUser) Then
                    'Es Caos el Dueño.
186                 If esCaos(attackerIndex) Then
                        'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
188                     Call WriteConsoleMsg(attackerIndex, "Los miembros de la Legión Oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
190                     PuedeAtacarNPC = False
                        Exit Function
                    End If
                End If
            End If
        End If
        
        'Es el Rey Preatoriano?
192     If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
194         If Not ClanPretoriano(Npclist(NpcIndex).ClanIndex).CanAtackMember(NpcIndex) Then
196             Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
    
            End If
    
        End If

198     PuedeAtacarNPC = True

        
        Exit Function

PuedeAtacarNPC_Err:
200     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.PuedeAtacarNPC", Erl)
202     Resume Next
        
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        
        On Error GoTo CalcularDarExp_Err
        

100     If UserList(UserIndex).Grupo.EnGrupo Then
102         Call CalcularDarExpGrupal(UserIndex, NpcIndex, ElDaño)
        Else

            Dim ExpaDar As Long
    
            '[Nacho] Chekeamos que las variables sean validas para las operaciones
104         If ElDaño <= 0 Then ElDaño = 0
106         If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
108         If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
    
            '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
110         ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

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
            
142             Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, ExpaDar, 6)

            End If

        End If

        
        Exit Sub

CalcularDarExp_Err:
144     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDarExp", Erl)
146     Resume Next
        
End Sub

Sub CalcularDarExpGrupal(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
        
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
100     If ElDaño <= 0 Then ElDaño = 0
102     If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
104     If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
    
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
106     ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

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

138         Case Else
140             BonificacionGrupo = 2
                
        End Select
 
142     If ExpMult > 0 Then
144         ExpaDar = ExpaDar * ExpMult
        
        End If
    
        Dim expbackup As Long

148     ExpaDar = ExpaDar * BonificacionGrupo

        Dim i     As Byte

        Dim index As Byte

152     ExpaDar = ExpaDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
        Dim ExpUser As Long
    
        If ExpaDar > 0 Then
154         For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
156             index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)
    
158             If UserList(index).flags.Muerto = 0 Then
160                 If Distancia(UserList(UserIndex).Pos, UserList(index).Pos) < 20 Then
162
164                     ExpUser = 0

166                     If UserList(index).donador.activo = 1 Then
168                         ExpUser = ExpaDar * 1.1
                        Else
170                         ExpUser = ExpaDar
                        End If
                    
172                     ExpUser = ExpUser * UserList(index).flags.ScrollExp
                
174                     If UserList(index).Stats.ELV < STAT_MAXELV Then
176                         UserList(index).Stats.Exp = UserList(index).Stats.Exp + ExpUser

178                         If UserList(index).Stats.Exp > MAXEXP Then UserList(index).Stats.Exp = MAXEXP

180                         If UserList(index).ChatCombate = 1 Then
182                             Call WriteLocaleMsg(index, "141", FontTypeNames.FONTTYPE_EXP, ExpUser)

                            End If

184                         Call WriteUpdateExp(index)
186                         Call CheckUserLevel(index)

                        End If
    
                    Else
    
                        'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
188                     If UserList(index).ChatCombate = 1 Then
190                         Call WriteLocaleMsg(index, "69", FontTypeNames.FONTTYPE_New_GRUPO)
    
                        End If
    
                    End If
    
                Else
    
208                 If UserList(index).ChatCombate = 1 Then
210                     Call WriteConsoleMsg(index, "Estás muerto, no has ganado experencia del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
    
                    End If
    
                End If
    
228         Next i
        End If

        'Else
        '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, experencia perdida.", FontTypeNames.FONTTYPE_New_GRUPO)
        'End If

        
        Exit Sub

CalcularDarExpGrupal_Err:
230     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDarExpGrupal", Erl)
232     Resume Next
        
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

        Dim index As Byte

132     OroDar = OroDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

134     For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
136         index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)

138         If UserList(index).flags.Muerto = 0 Then
140             If UserList(UserIndex).Pos.Map = UserList(index).Pos.Map Then
142                 If OroDar > 0 Then
                    
                        'OroDar = orobackup * UserList(Index).flags.ScrollOro
                
144                     UserList(index).Stats.GLD = UserList(index).Stats.GLD + OroDar
                        
146                     If UserList(index).ChatCombate = 1 Then
148                         Call WriteConsoleMsg(index, "¡El grupo ha ganado " & PonerPuntos(OroDar) & " monedas de oro!", FontTypeNames.FONTTYPE_New_GRUPO)

                        End If
                        
150                     Call WriteUpdateGold(index)
                        
                    End If

                Else

                    'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
                    'Call WriteLocaleMsg(Index, "69", FontTypeNames.FONTTYPE_INFOIAO)
                End If

            Else

                '  Call WriteConsoleMsg(Index, "Estas muerto, no has ganado oro del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
            End If

152     Next i

        'Else
        '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, oro perdido.", FontTypeNames.FONTTYPE_New_GRUPO)
        'End If

        
        Exit Sub

CalcularDarOroGrupal_Err:
154     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CalcularDarOroGrupal", Erl)
156     Resume Next
        
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

        'TODO: Pero que rebuscado!!
        'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
        On Error GoTo ErrHandler

        Dim tOrg As eTrigger

        Dim tDst As eTrigger
    
100     tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
102     tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
104     If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
106         If tOrg = tDst Then
108             TriggerZonaPelea = TRIGGER6_PERMITE
            Else
110             TriggerZonaPelea = TRIGGER6_PROHIBE

            End If

        Else
112         TriggerZonaPelea = TRIGGER6_AUSENTE

        End If

        Exit Function
ErrHandler:
114     TriggerZonaPelea = TRIGGER6_AUSENTE
116     LogError ("Error en TriggerZonaPelea - " & Err.description)

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
126     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserIncinera", Erl)
128     Resume Next
        
End Sub

Sub UserDañoEspecial(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
        
        On Error GoTo UserDañoEspecial_Err
        

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

UserDañoEspecial_Err:
198     Call RegistrarError(Err.Number, Err.description, "SistemaCombate.UserDañoEspecial", Erl)
200     Resume Next
        
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
        'Reaccion de las mascotas
        
        On Error GoTo AllMascotasAtacanUser_Err
    
        
        Dim iCount As Integer
    
100     For iCount = 1 To MAXMASCOTAS
102         If UserList(Maestro).MascotasIndex(iCount) > 0 Then
104             Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).name
106             Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
108             Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
            End If
110     Next iCount
        
        Exit Sub

AllMascotasAtacanUser_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.AllMascotasAtacanUser", Erl)

        
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
        
        On Error GoTo CheckPets_Err
    
        
        Dim j As Integer
    
100     For j = 1 To MAXMASCOTAS
102         If UserList(UserIndex).MascotasIndex(j) > 0 Then
104            If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
106             If CheckElementales Or (Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALAGUA And Npclist(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALVIENTO) Then
108                 If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
110                 Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                End If
               End If
            End If
112     Next j
        
        Exit Sub

CheckPets_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.CheckPets", Erl)

        
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
        
        On Error GoTo AllFollowAmo_Err
    
        
        Dim j As Integer
    
100     For j = 1 To MAXMASCOTAS
102         If UserList(UserIndex).MascotasIndex(j) > 0 Then
104             Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
            End If
106     Next j
        
        Exit Sub

AllFollowAmo_Err:
        Call RegistrarError(Err.Number, Err.description, "SistemaCombate.AllFollowAmo", Erl)

        
End Sub
