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


Private Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueArmas_Err
        

100     ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

        
        Exit Function

ModificadorPoderAtaqueArmas_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueArmas", Erl)
104     Resume Next
        
End Function

Private Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueProyectiles_Err
        
    
100     ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

        
        Exit Function

ModificadorPoderAtaqueProyectiles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueProyectiles", Erl)
104     Resume Next
        
End Function

Private Function ModicadorDañoClaseArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseArmas_Err
        
    
100     ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

        
        Exit Function

ModicadorDañoClaseArmas_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseArmas", Erl)
104     Resume Next
        
End Function

Private Function ModicadorApuñalarClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorApuñalarClase_Err
        
    
100     ModicadorApuñalarClase = ModClase(clase).ModApuñalar

        
        Exit Function

ModicadorApuñalarClase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorApuñalarClase", Erl)
104     Resume Next
        
End Function

Private Function ModicadorDañoClaseProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseProyectiles_Err
        
        
100     ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

        
        Exit Function

ModicadorDañoClaseProyectiles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseProyectiles", Erl)
104     Resume Next
        
End Function

Private Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModEvasionDeEscudoClase_Err
        

100     ModEvasionDeEscudoClase = ModClase(clase).Escudo

        
        Exit Function

ModEvasionDeEscudoClase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModEvasionDeEscudoClase", Erl)
104     Resume Next
        
End Function

Private Function Minimo(ByVal a As Single, ByVal b As Single) As Single
        
        On Error GoTo Minimo_Err
        

100     If a > b Then
102         Minimo = b
            Else:
104         Minimo = a

        End If

        
        Exit Function

Minimo_Err:
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.Minimo", Erl)
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
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.MinimoInt", Erl)
108     Resume Next
        
End Function

Private Function Maximo(ByVal a As Single, ByVal b As Single) As Single
        
        On Error GoTo Maximo_Err
        

100     If a > b Then
102         Maximo = a
            Else:
104         Maximo = b

        End If

        
        Exit Function

Maximo_Err:
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.Maximo", Erl)
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
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.MaximoInt", Erl)
108     Resume Next
        
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasionEscudo_Err
        

100     PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

        
        Exit Function

PoderEvasionEscudo_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderEvasionEscudo", Erl)
104     Resume Next
        
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasion_Err
        

        Dim lTemp As Long

100     With UserList(UserIndex)
102         lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).Evasion
       
104         PoderEvasion = (lTemp + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))

        End With

        
        Exit Function

PoderEvasion_Err:
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderEvasion", Erl)
108     Resume Next
        
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
        
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
116     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueArma", Erl)
118     Resume Next
        
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
        
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
116     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueProyectil", Erl)
118     Resume Next
        
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
        
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
116     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueWrestling", Erl)
118     Resume Next
        
End Function

Private Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo UserImpactoNpc_Err

        Dim PoderAtaque As Long

        Dim Arma        As Integer

        Dim Proyectil   As Boolean

        Dim ProbExito   As Long

100     Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex

102     If Arma = 0 Then Proyectil = False Else Proyectil = ObjData(Arma).Proyectil = 1

104     If Arma > 0 Then 'Usando un arma
106         If Proyectil Then
108             PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Else
110             PoderAtaque = PoderAtaqueArma(UserIndex)

            End If

        Else 'Peleando con puños
112         PoderAtaque = PoderAtaqueWrestling(UserIndex)

        End If

114     ProbExito = Maximo(10, Minimo(90, 70 + ((PoderAtaque - NpcList(NpcIndex).PoderEvasion) * 0.1)))

116     UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

118     If UserImpactoNpc Then
            Call SubirSkillDeArmaActual(UserIndex)
        End If

        Exit Function

UserImpactoNpc_Err:
130     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UserImpactoNpc", Erl)
132     Resume Next
        
End Function

Private Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
        
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
102     NpcPoderAtaque = NpcList(NpcIndex).PoderAtaque
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
128                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

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
136     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcImpacto", Erl)
138     Resume Next
        
End Function

Private Function CalcularDaño(ByVal UserIndex As Integer) As Long

        ' Reescrita por WyroX - 16/01/2021

        On Error GoTo CalcularDaño_Err

        Dim DañoUsuario As Long, DañoArma As Long, DañoMaxArma As Long, ModifClase As Single

        With UserList(UserIndex)
        
            ' Daño base del usuario
            DañoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)

            ' Daño con arma
            If .Invent.WeaponEqpObjIndex > 0 Then
                Dim Arma As ObjData
                Arma = ObjData(.Invent.WeaponEqpObjIndex)
                
                ' Calculamos el daño del arma
                DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                ' Daño máximo del arma
                DañoMaxArma = Arma.MaxHit

                ' Si lanza proyectiles
                If Arma.Proyectil = 1 Then
                    ' Usamos el modificador correspondiente
                    ModifClase = ModicadorDañoClaseProyectiles(.clase)

                    ' Si requiere munición
                    If Arma.Municion = 1 And .Invent.MunicionEqpObjIndex > 0 Then
                        Dim Municion As ObjData
                        Municion = ObjData(.Invent.MunicionEqpObjIndex)
                        ' Agregamos el daño de la munición al daño del arma
                        DañoArma = DañoArma + RandomNumber(Municion.MinHIT, Municion.MaxHit)
                        DañoMaxArma = Arma.MaxHit + Municion.MaxHit
                    End If
                
                ' Arma melé
                Else
                    ' Usamos el modificador correspondiente
                    ModifClase = ModicadorDañoClaseArmas(.clase)
                End If
        
            ' Daño con puños
            Else
                ' Modificador de combate sin armas
                ModifClase = ModClase(.clase).DañoWrestling
            
                ' Si tiene nudillos o guantes
                If .Invent.NudilloSlot > 0 Then
                    Arma = ObjData(.Invent.NudilloObjIndex)
                    
                    ' Calculamos el daño del nudillo o guante
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    ' Daño máximo
                    DañoMaxArma = Arma.MaxHit
                End If
            End If

            ' Calculo del daño
            CalcularDaño = (3 * DañoArma + DañoMaxArma * 0.2 * Maximo(0, .Stats.UserAtributos(Fuerza) - 15) + DañoUsuario) * ModifClase
            
            ' El pirata navegando pega un 20% más
            If .clase = eClass.Pirat And .flags.Navegando = 1 Then
                CalcularDaño = CalcularDaño * 1.2
            End If
            
            ' Daño del barco
            If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
                CalcularDaño = CalcularDaño + RandomNumber(ObjData(.Invent.BarcoObjIndex).MinHIT, ObjData(.Invent.BarcoObjIndex).MaxHit)

            ' Daño de la montura
            ElseIf .flags.Montado = 1 And .Invent.MonturaObjIndex > 0 Then
                CalcularDaño = CalcularDaño + RandomNumber(ObjData(.Invent.MonturaObjIndex).MinHIT, ObjData(.Invent.MonturaObjIndex).MaxHit)
            End If

        End With
        
        Exit Function

CalcularDaño_Err:
     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CalcularDaño", Erl)
     Resume Next
        
End Function

Private Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

        ' Reescrito por WyroX - 16/01/2021
        
        On Error GoTo UserDañoNpc_Err

        With UserList(UserIndex)

            Dim Daño As Long, DañoBase As Long, DañoExtra As Long, Color As Long, DañoStr As String

100         If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex And NpcList(NpcIndex).NPCtype = DRAGON Then
                ' Espada MataDragones
102             DañoBase = NpcList(NpcIndex).Stats.MinHp + NpcList(NpcIndex).Stats.def
                ' La pierde una vez usada
104             Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
            Else
                ' Daño normal
106             DañoBase = CalcularDaño(UserIndex)

                ' NPC de pruebas
108             If NpcList(NpcIndex).NPCtype = DummyTarget Then
110                 NpcList(NpcIndex).Contadores.UltimoAtaque = 30
                End If
            End If
            
            ' Color por defecto rojo
111         Color = vbRed

            ' Defensa del NPC
112         Daño = DañoBase - NpcList(NpcIndex).Stats.def

114         If Daño < 0 Then Daño = 0

            ' Mostramos en consola el golpe
116         If .ChatCombate = 1 Then
118             Call WriteLocaleMsg(UserIndex, "382", FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(Daño))
            End If

            ' Golpe crítico
120         If PuedeGolpeCritico(UserIndex) Then
                ' Si acertó - Doble chance contra NPCs
122             If RandomNumber(1, 100) <= ProbabilidadGolpeCritico(UserIndex) * 1.5 Then
                    ' Daño del golpe crítico (usamos el daño base)
124                 DañoExtra = DañoBase * ModDañoGolpeCritico
                
                    ' Mostramos en consola el daño
126                 If .ChatCombate = 1 Then
128                     Call WriteLocaleMsg(UserIndex, "383", FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(DañoExtra))
                    End If

                    ' Color naranja
130                 Color = RGB(225, 165, 0)
                End If

            ' Apuñalar (le afecta la defensa)
132         ElseIf PuedeApuñalar(UserIndex) Then
                ' Si acertó - Doble chance contra NPCs
136             If RandomNumber(1, 100) <= ProbabilidadApuñalar(UserIndex) * 1.5 Then
                    ' Daño del apuñalamiento
138                 DañoExtra = Daño * ModicadorApuñalarClase(.clase)
                
                    ' Mostramos en consola el daño
140                 If .ChatCombate = 1 Then
142                     Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(DañoExtra))
                    End If

                    ' Color amarillo
144                 Color = vbYellow
                End If

                ' Sube skills en apuñalar
146             Call SubirSkill(UserIndex, Apuñalar)
            End If
            
148         If DañoExtra > 0 Then
150             Daño = Daño + DañoExtra

                DañoStr = PonerPuntos(Daño)
                
                ' Mostramos el daño total en consola
152             If .ChatCombate = 1 Then
154                 Call WriteLocaleMsg(UserIndex, "384", FontTypeNames.FONTTYPE_FIGHT, DañoStr)
                End If
                
156             DañoStr = "¡" & DañoStr & "!"
            Else
158             DañoStr = PonerPuntos(Daño)
            End If

            ' Daño sobre el tile
160         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageTextCharDrop(DañoStr, NpcList(NpcIndex).Char.CharIndex, Color))

            ' Experiencia
162         Call CalcularDarExp(UserIndex, NpcIndex, Daño)

            ' Restamos el daño al NPC
164         NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MinHp - Daño

            ' NPC de invasión
            If NpcList(NpcIndex).flags.InvasionIndex Then
                Call SumarScoreInvasion(NpcList(NpcIndex).flags.InvasionIndex, UserIndex, Daño)
            End If

            ' Muere el NPC
166         If NpcList(NpcIndex).Stats.MinHp <= 0 Then
                ' Drop items, respawn, etc.
170             Call MuereNpc(NpcIndex, UserIndex)
            End If

        End With
        
        Exit Sub

UserDañoNpc_Err:
172     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UserDañoNpc", Erl)
174     Resume Next
        
End Sub

Private Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo NpcDaño_Err
        

        Dim Daño As Integer, Lugar As Integer, absorbido As Integer

        Dim antdaño As Integer, defbarco As Integer

        Dim obj As ObjData
    
100     Daño = RandomNumber(NpcList(NpcIndex).Stats.MinHIT, NpcList(NpcIndex).Stats.MaxHit)
102     antdaño = Daño
    
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
            ' 1/6 de chances de que sea a la cabeza
            Case PartesCuerpo.bCabeza

                'Si tiene casco absorbe el golpe
                If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                    Dim Casco As ObjData
                    Casco = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
                    absorbido = absorbido + RandomNumber(Casco.MinDef, Casco.MaxDef)
                End If

            Case Else

                'Si tiene armadura absorbe el golpe
                If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Dim Armadura As ObjData
                    Armadura = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)
                    absorbido = absorbido + RandomNumber(Armadura.MinDef, Armadura.MaxDef)
                End If
                
                'Si tiene escudo absorbe el golpe
                If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                    Dim Escudo As ObjData
                    Escudo = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
                    absorbido = absorbido + RandomNumber(Escudo.MinDef, Escudo.MaxDef)
                End If

        End Select
        
        Daño = Daño - absorbido - defbarco - defMontura
        
        If Daño < 0 Then Daño = 0
    
152     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageTextCharDrop(PonerPuntos(Daño), UserList(UserIndex).Char.CharIndex, vbRed))

154     If UserList(UserIndex).ChatCombate = 1 Then
156         Call WriteNPCHitUser(UserIndex, Lugar, Daño)
        End If

158     If UserList(UserIndex).flags.Privilegios And PlayerType.user Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Daño
    
160     If UserList(UserIndex).flags.Meditando Then
162         If Daño > Fix(UserList(UserIndex).Stats.MinHp / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
164             UserList(UserIndex).flags.Meditando = False
166             UserList(UserIndex).Char.FX = 0
168             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
            End If

        End If
    
        'Muere el usuario
170     If UserList(UserIndex).Stats.MinHp <= 0 Then
    
172         Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
                    
178         If NpcList(NpcIndex).MaestroUser > 0 Then
180             Call AllFollowAmo(NpcList(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
184             NpcList(NpcIndex).Movement = NpcList(NpcIndex).flags.OldMovement
186             NpcList(NpcIndex).Hostile = NpcList(NpcIndex).flags.OldHostil
188             NpcList(NpcIndex).flags.AttackedBy = vbNullString
                NpcList(NpcIndex).Target = 0
            End If
        
190         Call UserDie(UserIndex)

        Else
192         Call WriteUpdateHP(UserIndex)
    
        End If

        
        Exit Sub

NpcDaño_Err:
194     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcDaño", Erl)
196     Resume Next
        
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Heading As eHeading) As Boolean
        
    On Error GoTo NpcAtacaUser_Err
        

    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.user) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
    ' El npc puede atacar ???
    
    If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
        NpcAtacaUser = False
        Exit Function
    End If
        
    If ((MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Blocked And 2 ^ (Heading - 1)) <> 0) Then
        NpcAtacaUser = False
        Exit Function
    End If

    NpcAtacaUser = True

    Call AllMascotasAtacanNPC(NpcIndex, UserIndex)

    If NpcList(NpcIndex).Target = 0 Then NpcList(NpcIndex).Target = UserIndex
    
    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
    
    If NpcList(NpcIndex).flags.Snd1 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd1, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
    End If
        
    Call CancelExit(UserIndex)

    If NpcImpacto(NpcIndex, UserIndex) Then
    
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        
        If UserList(UserIndex).flags.Navegando = 0 Or UserList(UserIndex).flags.Montado = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))
        End If
        
        Call NpcDaño(NpcIndex, UserIndex)

        '¿Puede envenenar?
        If NpcList(NpcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, NpcList(NpcIndex).Veneno)
        
    Else
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharSwing(NpcList(NpcIndex).Char.CharIndex, False))

    End If

    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
        

    Exit Function

NpcAtacaUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcAtacaUser", Erl)
    Resume Next
        
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
        
        On Error GoTo NpcImpactoNpc_Err
        

        Dim PoderAtt  As Long, PoderEva As Long

        Dim ProbExito As Long

100     PoderAtt = NpcList(Atacante).PoderAtaque
102     PoderEva = NpcList(Victima).PoderEvasion
104     ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtt - PoderEva) * 0.4)))
106     NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

        
        Exit Function

NpcImpactoNpc_Err:
108     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcImpactoNpc", Erl)
110     Resume Next
        
End Function

Private Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
        
            On Error GoTo NpcDañoNpc_Err

            Dim Daño As Integer
    
100         With NpcList(Atacante)
102             Daño = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
104             NpcList(Victima).Stats.MinHp = NpcList(Victima).Stats.MinHp - Daño
            
106             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessageTextCharDrop(PonerPuntos(Daño), NpcList(Victima).Char.CharIndex, vbRed))
            
                ' Mascotas dan experiencia al amo
108             If .MaestroUser > 0 Then
110                 Call CalcularDarExp(.MaestroUser, Victima, Daño)
                End If
            
112             If NpcList(Victima).Stats.MinHp < 1 Then
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
126         Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcDañoNpc")
128         Resume Next
        
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMovimiento As Boolean = True)
        
        On Error GoTo NpcAtacaNpc_Err
        
        If Not IntervaloPermiteAtacarNPC(Atacante) Then Exit Sub
        
100     If cambiarMovimiento Then
106         NpcList(Victima).TargetNPC = Atacante
108         NpcList(Victima).Movement = TipoAI.NpcAtacaNpc
        End If

110     If NpcList(Atacante).flags.Snd1 > 0 Then
112         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(NpcList(Atacante).flags.Snd1, NpcList(Atacante).pos.x, NpcList(Atacante).pos.y))

        End If

114     If NpcImpactoNpc(Atacante, Victima) Then
    
116         If NpcList(Victima).flags.Snd2 > 0 Then
118             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(NpcList(Victima).flags.Snd2, NpcList(Victima).pos.x, NpcList(Victima).pos.y))
            Else
120             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(Victima).pos.x, NpcList(Victima).pos.y))

            End If

122         Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, NpcList(Victima).pos.x, NpcList(Victima).pos.y))
    
124         Call NpcDañoNpc(Atacante, Victima)
    
        Else
126         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageCharSwing(NpcList(Atacante).Char.CharIndex, False, True))

        End If

        
        Exit Sub

NpcAtacaNpc_Err:
128     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcAtacaNpc", Erl)
130     Resume Next
        
End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        
        On Error GoTo UsuarioAtacaNpc_Err
        
        If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then Exit Sub

        If UserList(UserIndex).flags.invisible = 0 Then Call NPCAtacado(NpcIndex, UserIndex)

        If UserImpactoNpc(UserIndex, NpcIndex) Then
        
            ' Suena el Golpe en el cliente.
            If NpcList(NpcIndex).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
            End If
        
            ' Golpe Paralizador
            If UserList(UserIndex).flags.Paraliza = 1 And NpcList(NpcIndex).flags.Paralizado = 0 Then

                If RandomNumber(1, 4) = 1 Then

                    If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
                        NpcList(NpcIndex).flags.Paralizado = 1
                        NpcList(NpcIndex).Contadores.Paralisis = (IntervaloParalizado / 3) * 7

                        If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
                            Call WriteLocaleMsg(UserIndex, "136", FontTypeNames.FONTTYPE_FIGHT)

                        End If

                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.CharIndex, 8, 0))
                                 
                    Else

                        If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                End If

            End If
            
            ' Cambiamos el objetivo del NPC si uno le pega cuerpo a cuerpo.
            If NpcList(NpcIndex).Target <> UserIndex Then
                NpcList(NpcIndex).Target = UserIndex
            End If
            
            ' Si te mimetizaste en forma de bicho y le pegas al chobi, el chobi te va a pegar.
            If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.FormaBicho Then
                UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.FormaBichoSinProteccion
            End If
            
            ' Resta la vida del NPC
136         Call UserDañoNpc(UserIndex, NpcIndex)
            
            Dim Arma As Integer: Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
            Dim municionIndex As Integer: municionIndex = UserList(UserIndex).Invent.MunicionEqpObjIndex
            Dim Particula As Integer
            Dim Tiempo    As Long
            
            If Arma > 0 Then
                If municionIndex > 0 And ObjData(Arma).Proyectil Then
                    If ObjData(municionIndex).CreaFX <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.CharIndex, ObjData(municionIndex).CreaFX, 0))
                    
                    End If
                                        
                    If ObjData(municionIndex).CreaParticula <> "" Then
                        Particula = val(ReadField(1, ObjData(municionIndex).CreaParticula, Asc(":")))
                        Tiempo = val(ReadField(2, ObjData(municionIndex).CreaParticula, Asc(":")))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(NpcList(NpcIndex).Char.CharIndex, Particula, Tiempo, False))
                    End If
                End If
            End If
            
        Else
            
            Dim sendto As SendTarget
            
            If UserList(UserIndex).clase = eClass.Hunter And UserList(UserIndex).flags.Oculto = 0 Then
                sendto = SendTarget.ToPCArea
            Else
                sendto = SendTarget.ToIndex
            End If

            Call SendData(sendto, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))

        End If

        
        Exit Sub

UsuarioAtacaNpc_Err:
146     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaNpc", Erl)
148     Resume Next
        
End Sub

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

        Dim AttackPos As WorldPos
122         AttackPos = UserList(UserIndex).pos

124     Call HeadtoPos(UserList(UserIndex).Char.Heading, AttackPos)
       
        'Exit if not legal
126     If AttackPos.x >= XMinMapSize And AttackPos.x <= XMaxMapSize And AttackPos.y >= YMinMapSize And AttackPos.y <= YMaxMapSize Then

128         If ((MapData(AttackPos.Map, AttackPos.x, AttackPos.y).Blocked And 2 ^ (UserList(UserIndex).Char.Heading - 1)) <> 0) Then
130             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))
                Exit Sub
            End If

            Dim index As Integer

132         index = MapData(AttackPos.Map, AttackPos.x, AttackPos.y).UserIndex

            'Look for user
134         If index > 0 Then
136             Call UsuarioAtacaUsuario(UserIndex, index)

            'Look for NPC
142         ElseIf MapData(AttackPos.Map, AttackPos.x, AttackPos.y).NpcIndex > 0 Then

144             index = MapData(AttackPos.Map, AttackPos.x, AttackPos.y).NpcIndex

146             If NpcList(index).Attackable Then
148                 If NpcList(index).MaestroUser > 0 And MapInfo(NpcList(index).pos.Map).Seguro = 1 Then
150                     Call WriteConsoleMsg(UserIndex, "No podés atacar mascotas en zonas seguras", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If

152                 Call UsuarioAtacaNpc(UserIndex, index)

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
162     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtaca", Erl)
164     Resume Next
        
End Sub

Private Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean

        On Error GoTo UsuarioImpacto_Err

        Dim ProbRechazo            As Long
        Dim Rechazo                As Boolean
        Dim ProbExito              As Long
        Dim PoderAtaque            As Long
        Dim UserPoderEvasion       As Long
        Dim Arma                   As Integer
        Dim Proyectil              As Boolean
        Dim SkillTacticas          As Long
        Dim SkillDefensa           As Long

100     If UserList(AtacanteIndex).flags.GolpeCertero = 1 Then
102         UsuarioImpacto = True
104         UserList(AtacanteIndex).flags.GolpeCertero = 0
            Exit Function

        End If

106     SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
108     SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)

110     Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex

112     If Arma > 0 Then
            Proyectil = ObjData(Arma).Proyectil = 1

            If Proyectil Then
                PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
            Else
                PoderAtaque = PoderAtaqueArma(AtacanteIndex)
            End If
        Else
116         Proyectil = False
            PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
        End If

        'Calculamos el poder de evasion...
118     UserPoderEvasion = PoderEvasion(VictimaIndex)

        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
            UserPoderEvasion = UserPoderEvasion + PoderEvasionEscudo(VictimaIndex)
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
        Else
            ProbRechazo = 0
        End If

        ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))

        UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)

        If UsuarioImpacto Then
          Call SubirSkillDeArmaActual(AtacanteIndex)

        Else ' Falló
            If RandomNumber(1, 100) <= ProbRechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageEscudoMov(UserList(VictimaIndex).Char.CharIndex))

                If UserList(AtacanteIndex).ChatCombate = 1 Then
                    Call WriteBlockedWithShieldOther(AtacanteIndex)
                End If

                If UserList(VictimaIndex).ChatCombate = 1 Then
                    Call WriteBlockedWithShieldUser(VictimaIndex)
                End If

                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 88, 0))
                Call SubirSkill(VictimaIndex, eSkill.Defensa)
            Else
                Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te atacó y falló! ", FontTypeNames.FONTTYPE_FIGHT)

            End If
        End If

        Exit Function

UsuarioImpacto_Err:
        Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioImpacto", Erl)
        Resume Next

End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
        
        On Error GoTo UsuarioAtacaUsuario_Err

        Dim sendto As SendTarget
        Dim Probabilidad As Byte
        Dim HuboEfecto   As Boolean
            HuboEfecto = False

100     If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub

102     If Distancia(UserList(AtacanteIndex).pos, UserList(VictimaIndex).pos) > MAXDISTANCIAARCO Then
104         Call WriteLocaleMsg(AtacanteIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(atacanteindex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

108     Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

110     If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then

114         Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))

116         If UserList(VictimaIndex).flags.Navegando = 0 Or UserList(VictimaIndex).flags.Montado = 0 Then
118             Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If

128         Call UserDañoUser(AtacanteIndex, VictimaIndex)

        Else

130         If UserList(AtacanteIndex).flags.invisible Or UserList(AtacanteIndex).flags.Oculto Then
134             sendto = SendTarget.ToIndex
            Else
132             sendto = SendTarget.ToPCArea
            End If

136         Call SendData(sendto, AtacanteIndex, PrepareMessageCharSwing(UserList(AtacanteIndex).Char.CharIndex))

        End If

        Exit Sub

UsuarioAtacaUsuario_Err:
138     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaUsuario", Erl)
140     Resume Next

End Sub

Private Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
        On Error GoTo UserDañoUser_Err

100     With UserList(VictimaIndex)

            Dim Daño As Long, DañoBase As Long, DañoExtra As Long, Defensa As Long, Color As Long, DañoStr As String, Lugar As PartesCuerpo

            ' Daño normal
102         DañoBase = CalcularDaño(AtacanteIndex)

            ' Color por defecto rojo
103         Color = vbRed

            ' Elegimos al azar una parte del cuerpo
104         Lugar = RandomNumber(1, 6)

            Select Case Lugar
                ' 1/6 de chances de que sea a la cabeza
                Case PartesCuerpo.bCabeza

                    'Si tiene casco absorbe el golpe
106                 If .Invent.CascoEqpObjIndex > 0 Then
                        Dim Casco As ObjData
108                     Casco = ObjData(.Invent.CascoEqpObjIndex)
110                     Defensa = Defensa + RandomNumber(Casco.MinDef, Casco.MaxDef)
                    End If

                Case Else

                    'Si tiene armadura absorbe el golpe
112                 If .Invent.ArmourEqpObjIndex > 0 Then
                        Dim Armadura As ObjData
114                     Armadura = ObjData(.Invent.ArmourEqpObjIndex)
116                     Defensa = Defensa + RandomNumber(Armadura.MinDef, Armadura.MaxDef)
                    End If
                    
                    'Si tiene escudo absorbe el golpe
118                 If .Invent.EscudoEqpObjIndex > 0 Then
                        Dim Escudo As ObjData
120                     Escudo = ObjData(.Invent.EscudoEqpObjIndex)
122                     Defensa = Defensa + RandomNumber(Escudo.MinDef, Escudo.MaxDef)
                    End If
    
            End Select

            ' Defensa del barco de la víctima
124         If .Invent.BarcoObjIndex > 0 Then
                Dim Barco As ObjData
126             Barco = ObjData(.Invent.BarcoObjIndex)
128             Defensa = Defensa + RandomNumber(Barco.MinDef, Barco.MaxDef)

            ' Defensa de la montura de la víctima
130         ElseIf .Invent.MonturaObjIndex > 0 Then
                Dim Montura As ObjData
132             Montura = ObjData(.Invent.MonturaObjIndex)
134             Defensa = Defensa + RandomNumber(Montura.MinDef, Montura.MaxDef)
            End If
            
            ' Refuerzo de la espada - Ignora parte de la armadura
136         If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
138             Defensa = Defensa - ObjData(UserList(AtacanteIndex).Invent.WeaponEqpObjIndex).Refuerzo

140             If Defensa < 0 Then Defensa = 0
            End If
            
            ' Restamos la defensa
142         Daño = DañoBase - Defensa

144         If Daño < 0 Then Daño = 0

            DañoStr = PonerPuntos(Daño)

            ' Mostramos en consola el golpe al atacante
146         If UserList(AtacanteIndex).ChatCombate = 1 Then
148             Call WriteUserHittedUser(AtacanteIndex, Lugar, .Char.CharIndex, DañoStr)
            End If
            ' Y a la víctima
150         If .ChatCombate = 1 Then
152             Call WriteUserHittedByUser(VictimaIndex, Lugar, UserList(AtacanteIndex).Char.CharIndex, DañoStr)
            End If

            ' Golpe crítico (ignora defensa)
154         If PuedeGolpeCritico(AtacanteIndex) Then
                ' Si acertó
156             If RandomNumber(1, 100) <= ProbabilidadGolpeCritico(AtacanteIndex) Then
                    ' Daño del golpe crítico (usamos el daño base)
158                 DañoExtra = DañoBase * ModDañoGolpeCritico

                    DañoStr = PonerPuntos(DañoExtra)

                    ' Mostramos en consola el daño al atacante
160                 If UserList(AtacanteIndex).ChatCombate = 1 Then
162                     Call WriteLocaleMsg(AtacanteIndex, "383", FontTypeNames.FONTTYPE_FIGHT, .name & "¬" & DañoStr)
                    End If
                    ' Y a la víctima
164                 If .ChatCombate = 1 Then
166                     Call WriteLocaleMsg(VictimaIndex, "385", FontTypeNames.FONTTYPE_FIGHT, UserList(AtacanteIndex).name & "¬" & DañoStr)
                    End If

                    ' Color naranja
168                 Color = RGB(225, 165, 0)
                End If

            ' Apuñalar (le afecta la defensa)
170         ElseIf PuedeApuñalar(AtacanteIndex) Then
172             If RandomNumber(1, 100) <= ProbabilidadApuñalar(AtacanteIndex) Then
                    ' Daño del apuñalamiento
174                 DañoExtra = Daño * ModicadorApuñalarClase(UserList(AtacanteIndex).clase)

                    DañoStr = PonerPuntos(DañoExtra)
                
                    ' Mostramos en consola el daño al atacante
176                 If UserList(AtacanteIndex).ChatCombate = 1 Then
178                     Call WriteLocaleMsg(AtacanteIndex, "210", FontTypeNames.FONTTYPE_FIGHT, .name & "¬" & DañoStr)
                    End If
                    ' Y a la víctima
180                 If .ChatCombate = 1 Then
182                     Call WriteLocaleMsg(VictimaIndex, "211", FontTypeNames.FONTTYPE_FIGHT, UserList(AtacanteIndex).name & "¬" & DañoStr)
                    End If

                    ' Color amarillo
184                 Color = vbYellow

                    ' Efecto en la víctima
186                 Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 89, 0))
                    
                    ' Efecto en pantalla a ambos
188                 Call WriteFlashScreen(VictimaIndex, &H3C3CFF, 200, True)
190                 Call WriteFlashScreen(AtacanteIndex, &H3C3CFF, 150, True)
                End If

                ' Sube skills en apuñalar
192             Call SubirSkill(AtacanteIndex, Apuñalar)
            ElseIf PuedeDesequiparDeUnGolpe(AtacanteIndex) Then
                If RandomNumber(1, 100) <= ProbabilidadDesequipar(AtacanteIndex) Then
                    Call DesequiparObjetoDeUnGolpe(AtacanteIndex, VictimaIndex, Lugar)
                End If

            End If
            
196         If DañoExtra > 0 Then
198             Daño = Daño + DañoExtra

                DañoStr = PonerPuntos(Daño)
                
                ' Mostramos el daño total en consola al atacante
200             If UserList(AtacanteIndex).ChatCombate = 1 Then
202                 Call WriteLocaleMsg(AtacanteIndex, "384", FontTypeNames.FONTTYPE_FIGHT, DañoStr)
                End If
                ' Y a la víctima
204             If .ChatCombate = 1 Then
206                 Call WriteLocaleMsg(VictimaIndex, "387", FontTypeNames.FONTTYPE_FIGHT, UserList(AtacanteIndex).name & "¬" & DañoStr)
                End If
                
208             DañoStr = "¡" & PonerPuntos(Daño) & "!"
            Else
210             DañoStr = PonerPuntos(Daño)
            End If

            ' Daño sobre el tile
212         Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageTextCharDrop(DañoStr, .Char.CharIndex, Color))

            ' Restamos el daño a la víctima
214         .Stats.MinHp = .Stats.MinHp - Daño

            ' Muere la víctima
216         If .Stats.MinHp <= 0 Then
                ' Sumar frag y rutina de muerte
218             Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
220             Call ContarMuerte(VictimaIndex, AtacanteIndex)
222             Call ActStats(VictimaIndex, AtacanteIndex)
            ' Si sigue vivo
            Else
                ' Enviamos la vida
224             Call WriteUpdateHP(VictimaIndex)

                ' Intentamos aplicar algún efecto de estado
226             Call UserDañoEspecial(AtacanteIndex, VictimaIndex)
            End If

        End With

        Exit Sub

UserDañoUser_Err:
228     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UserDañoUser", Erl)
230     Resume Next
        
End Sub

Private Sub DesequiparObjetoDeUnGolpe(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer, ByVal parteDelCuerpo As PartesCuerpo)
    On Error GoTo DesequiparObjetoDeUnGolpe_Err
    
    Dim desequiparCasco As Boolean, desequiparArma As Boolean, desequiparEscudo As Boolean
    
    With UserList(VictimIndex)
    
        Select Case parteDelCuerpo
        Case PartesCuerpo.bCabeza
            ' Si pega en la cabeza, desequipamos el casco si tiene
            desequiparCasco = .Invent.CascoEqpObjIndex > 0
            ' Si no tiene casco, intentaremos desequipar otra cosa porque un golpe en la cabeza
            ' algo te tiene que desequipar.
            desequiparArma = (Not desequiparCasco) And (.Invent.WeaponEqpObjIndex > 0)
            desequiparEscudo = (Not desequiparCasco) And (Not desequiparArma) And (.Invent.EscudoEqpObjIndex > 0)
         
        Case PartesCuerpo.bBrazoDerecho, PartesCuerpo.bBrazoIzquierdo, PartesCuerpo.bTorso
            desequiparArma = (.Invent.WeaponEqpObjIndex > 0)
            desequiparEscudo = (Not desequiparArma) And (.Invent.EscudoEqpObjIndex > 0)
            desequiparCasco = False
            
        Case PartesCuerpo.bPiernaDerecha, PartesCuerpo.bPiernaIzquierda
            desequiparEscudo = (.Invent.EscudoEqpObjIndex > 0)
            desequiparCasco = False
            desequiparArma = False
            
        End Select
        
        If desequiparCasco Then
            Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
            
            Call WriteCombatConsoleMsg(AttackerIndex, "Has logrado desequipar el casco de tu oponente!")
            Call WriteCombatConsoleMsg(VictimIndex, UserList(AttackerIndex).name & " te ha desequipado el casco.")
            
        ElseIf desequiparArma Then
            Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
            Call WriteCombatConsoleMsg(AttackerIndex, "Has logrado desarmar a tu oponente!")
            Call WriteCombatConsoleMsg(VictimIndex, UserList(AttackerIndex).name & " te ha desarmado.")

        ElseIf desequiparEscudo Then
            Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
            Call WriteCombatConsoleMsg(AttackerIndex, "Has logrado desequipar el escudo de " & .name & ".")
            Call WriteCombatConsoleMsg(VictimIndex, UserList(AttackerIndex).name & " te ha desequipado el escudo.")
        Else
            Call WriteCombatConsoleMsg(AttackerIndex, "No has logrado desequipar ningun item a tu oponente!")
        End If
            
    End With
        
    Exit Sub

DesequiparObjetoDeUnGolpe_Err:
    Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.DesequiparObjetoDeUnGolpe", Erl)
 
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
        '***************************************************
        'Autor: Unknown
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************

        On Error GoTo UsuarioAtacadoPorUsuario_Err

        'Si la victima esta saliendo se cancela la salida
        Call CancelExit(VictimIndex)

100     If UserList(VictimIndex).flags.Meditando Then
102         UserList(VictimIndex).flags.Meditando = False
104         UserList(VictimIndex).Char.FX = 0
106         Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.CharIndex, 0))
        End If
    
108     If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
        Dim EraCriminal As Byte
    
110     UserList(VictimIndex).Counters.TiempoDeMapeo = 3
112     UserList(AttackerIndex).Counters.TiempoDeMapeo = 3
    
114     If Status(AttackerIndex) = 1 And Status(VictimIndex) = 1 Or Status(VictimIndex) = 3 Then
116         Call VolverCriminal(AttackerIndex)

        End If

118     EraCriminal = Status(AttackerIndex)
    
120     If EraCriminal = 2 And Status(AttackerIndex) < 2 Then
122         Call RefreshCharStatus(AttackerIndex)
124     ElseIf EraCriminal < 2 And Status(AttackerIndex) = 2 Then
126         Call RefreshCharStatus(AttackerIndex)
        End If

128     If Status(AttackerIndex) = 2 Then If UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(AttackerIndex)


130     Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
132     Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)

        Exit Sub

UsuarioAtacadoPorUsuario_Err:
136     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacadoPorUsuario", Erl)
138     Resume Next
        
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
        
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
100     If UserList(AttackerIndex).flags.Muerto = 1 Then
102         Call WriteLocaleMsg(AttackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacar = False
            Exit Function

        End If
        
        If UserList(AttackerIndex).flags.EnReto Then
            If Retos.Salas(UserList(AttackerIndex).flags.SalaReto).TiempoItems > 0 Then
                Call WriteConsoleMsg(AttackerIndex, "No podés atacar en este momento.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacar = False
                Exit Function
            End If
        End If

        'No podes atacar a alguien muerto
106     If UserList(VictimIndex).flags.Muerto = 1 Then
108         Call WriteConsoleMsg(AttackerIndex, "No podés atacar a un espiritu.", FontTypeNames.FONTTYPE_INFO)
110         PuedeAtacar = False
            Exit Function

        End If
        
        ' No podes atacar si estas en consulta
112     If UserList(AttackerIndex).flags.EnConsulta Then
114         Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
    
        End If
        
        ' No podes atacar si esta en consulta
116     If UserList(VictimIndex).flags.EnConsulta Then
118         Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
    
        End If
        
120     If UserList(AttackerIndex).flags.Maldicion = 1 Then
122         Call WriteConsoleMsg(AttackerIndex, "¡Estas maldito! No podes atacar.", FontTypeNames.FONTTYPE_INFO)
124         PuedeAtacar = False
            Exit Function

        End If
        
        If UserList(AttackerIndex).flags.Montado = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "No podés atacar usando una montura.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function

        End If

        'Estamos en una Arena? o un trigger zona segura?
138     T = TriggerZonaPelea(AttackerIndex, VictimIndex)

140     If T = eTrigger6.TRIGGER6_PERMITE Then
142         PuedeAtacar = True
            Exit Function
144     ElseIf T = eTrigger6.TRIGGER6_PROHIBE Then
146         PuedeAtacar = False
            Exit Function
148     ElseIf T = eTrigger6.TRIGGER6_AUSENTE Then

            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            ' If Not UserList(VictimIndex).flags.Privilegios And PlayerType.User Then
            '   If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
            ' PuedeAtacar = False
            '    Exit Function
            ' End If
        End If
        
        'Solo administradores pueden atacar a usuarios (PARA TESTING)
160     If (UserList(AttackerIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Admin)) = 0 Then
162         PuedeAtacar = False
            Exit Function
        End If
        
        'Estas queriendo atacar a un GM?
164     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

166     If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(AttackerIndex).flags.Privilegios And rank) Then
168         If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
170         PuedeAtacar = False
            Exit Function

        End If

        'Sos un Armada atacando un ciudadano?
172     If (Status(VictimIndex) = 1) And (esArmada(AttackerIndex)) Then
174         Call WriteConsoleMsg(AttackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
176         PuedeAtacar = False
            Exit Function

        End If

        'Tenes puesto el seguro?
178     If UserList(AttackerIndex).flags.Seguro Then
180         If Status(VictimIndex) = 1 Then
182             Call WriteConsoleMsg(AttackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
184             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Es un ciuda queriando atacar un imperial?
186     If UserList(AttackerIndex).flags.Seguro Then
188         If (Status(AttackerIndex) = 1) And (esArmada(VictimIndex)) Then
190             Call WriteConsoleMsg(AttackerIndex, "Los ciudadanos no pueden atacar a los soldados imperiales.", FontTypeNames.FONTTYPE_WARNING)
192             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Estas en un Mapa Seguro?
194     If MapInfo(UserList(VictimIndex).pos.Map).Seguro = 1 Then

196         If esArmada(AttackerIndex) Then
198             If UserList(AttackerIndex).Faccion.RecompensasReal >= 3 Then
200                 If UserList(VictimIndex).pos.Map = 58 Or UserList(VictimIndex).pos.Map = 59 Or UserList(VictimIndex).pos.Map = 60 Then
202                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
204                     PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

206         If esCaos(AttackerIndex) Then
208             If UserList(AttackerIndex).Faccion.RecompensasCaos >= 3 Then
210                 If UserList(VictimIndex).pos.Map = 195 Or UserList(VictimIndex).pos.Map = 196 Then
212                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
214                     PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

216         Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
218         PuedeAtacar = False
            Exit Function

        End If

        'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
220     If MapData(UserList(VictimIndex).pos.Map, UserList(VictimIndex).pos.x, UserList(VictimIndex).pos.y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(AttackerIndex).pos.Map, UserList(AttackerIndex).pos.x, UserList(AttackerIndex).pos.y).trigger = eTrigger.ZONASEGURA Then
222         Call WriteConsoleMsg(AttackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
224         PuedeAtacar = False
            Exit Function

        End If

226     PuedeAtacar = True

        
        Exit Function

PuedeAtacar_Err:
228     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeAtacar", Erl)
230     Resume Next
        
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
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
100     If UserList(AttackerIndex).flags.Muerto = 1 Then
            'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
102         Call WriteLocaleMsg(AttackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
104         PuedeAtacarNPC = False
            Exit Function

        End If
             
        If UserList(AttackerIndex).flags.Montado = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "No podés atacar usando una montura.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function

        End If

        'Solo administradores pueden atacar a usuarios (PARA TESTING)
106     If (UserList(AttackerIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) = 0 Then
108         PuedeAtacarNPC = False
            Exit Function
        End If
        
        ' No podes atacar si estas en consulta
110     If UserList(AttackerIndex).flags.EnConsulta Then
112         Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function

        End If
        
        'Es una criatura atacable?
114     If NpcList(NpcIndex).Attackable = 0 Then
            'No es una criatura atacable
116         Call WriteConsoleMsg(AttackerIndex, "No podés atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
118         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es valida la distancia a la cual estamos atacando?
120     If Distancia(UserList(AttackerIndex).pos, NpcList(NpcIndex).pos) >= MAXDISTANCIAARCO Then
122         Call WriteLocaleMsg(AttackerIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
124         PuedeAtacarNPC = False
            Exit Function

        End If


        'Si el usuario pertenece a una faccion
        If esArmada(AttackerIndex) Or esCaos(AttackerIndex) Then
            ' Y el NPC pertenece a la misma faccion
            If NpcList(NpcIndex).flags.Faccion = UserList(AttackerIndex).Faccion.Status Then
                Call WriteConsoleMsg(AttackerIndex, "No podés atacar NPCs de tu misma facción, para hacerlo debes desenlistarte.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
            
            ' Si es una mascota, checkeamos en el Maestro
            If NpcList(NpcIndex).MaestroUser > 0 Then
                If UserList(NpcList(NpcIndex).MaestroUser).Faccion.Status = UserList(AttackerIndex).Faccion.Status Then
                    Call WriteConsoleMsg(AttackerIndex, "No podés atacar NPCs de tu misma facción, para hacerlo debes desenlistarte.", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                End If
            End If
        End If
        
        If Status(AttackerIndex) = Ciudadano Then
            If NpcList(NpcIndex).MaestroUser > 0 And NpcList(NpcIndex).MaestroUser = AttackerIndex Then
                Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a tus mascotas siendo un ciudadano.", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            End If
        End If
        
        
        ' El seguro es SOLO para ciudadanos. La armada debe desenlistarse antes de querer atacar y se checkea arriba.
        ' Los criminales o Caos, ya estan mas alla del seguro.
        If Status(AttackerIndex) = Ciudadano Then
            
            If NpcList(NpcIndex).flags.Faccion = Armada Then
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Debes quitar el seguro para atacar miembros de la Armada Real (/seg)", FontTypeNames.FONTTYPE_INFO)
                    PuedeAtacarNPC = False
                    Exit Function
                Else
                    Call WriteConsoleMsg(AttackerIndex, "Atacaste un miembro de la Armada Real! Te has convertido en un Criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(AttackerIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                End If
            End If
            
            'Es el NPC mascota de alguien?
            If NpcList(NpcIndex).MaestroUser > 0 Then
                Select Case UserList(NpcList(NpcIndex).MaestroUser).Faccion.Status
                    Case e_Facciones.Armada
                        If UserList(AttackerIndex).flags.Seguro Then
                            Call WriteConsoleMsg(AttackerIndex, "Debes quitar el seguro para atacar mascotas de la Armada Real (/seg)", FontTypeNames.FONTTYPE_INFO)
                            PuedeAtacarNPC = False
                            Exit Function
                        Else
                            Call WriteConsoleMsg(AttackerIndex, "Atacaste una mascota de la Armada Real! Te has convertido en un Criminal.", FontTypeNames.FONTTYPE_INFO)
                            Call VolverCriminal(AttackerIndex)
                            PuedeAtacarNPC = True
                            Exit Function
                        End If
                        
                    Case e_Facciones.Ciudadano
                        If UserList(AttackerIndex).flags.Seguro Then
                            Call WriteConsoleMsg(AttackerIndex, "Debes quitar el seguro para atacar mascotas de otros ciudadanos(/seg)", FontTypeNames.FONTTYPE_INFO)
                            PuedeAtacarNPC = False
                            Exit Function
                        Else
                            Call WriteConsoleMsg(AttackerIndex, "Atacaste un la mascota de un ciudadano! Te has convertido en un Criminal.", FontTypeNames.FONTTYPE_INFO)
                            Call VolverCriminal(AttackerIndex)
                            PuedeAtacarNPC = True
                            Exit Function
                        End If
                    
                    Case Else
                        PuedeAtacarNPC = True
                        Exit Function
                End Select
            End If
        End If
        
        'Es el Rey Preatoriano?
196     If NpcList(NpcIndex).NPCtype = eNPCType.Pretoriano Then
198         If Not ClanPretoriano(NpcList(NpcIndex).ClanIndex).CanAtackMember(NpcIndex) Then
200             Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ejercito antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function
    
            End If
    
        End If

202     PuedeAtacarNPC = True

        
        Exit Function

PuedeAtacarNPC_Err:
204     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeAtacarNPC", Erl)
206     Resume Next
        
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        
        On Error GoTo CalcularDarExp_Err
        
        If NpcList(NpcIndex).MaestroUser <> 0 Then
            Exit Sub
        End If

100     If UserList(UserIndex).Grupo.EnGrupo Then
102         Call CalcularDarExpGrupal(UserIndex, NpcIndex, ElDaño)
        Else

            Dim ExpaDar As Double
    
            '[Nacho] Chekeamos que las variables sean validas para las operaciones
104         If ElDaño <= 0 Then ElDaño = 0
106         If NpcList(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub

            '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
            
110         ExpaDar = CDbl(ElDaño) * CDbl(NpcList(NpcIndex).GiveEXP) / NpcList(NpcIndex).Stats.MaxHp

112         If ExpaDar <= 0 Then Exit Sub

            '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
114         If ExpaDar > NpcList(NpcIndex).flags.ExpCount Then
116             ExpaDar = NpcList(NpcIndex).flags.ExpCount
118             NpcList(NpcIndex).flags.ExpCount = 0
            Else
120             NpcList(NpcIndex).flags.ExpCount = NpcList(NpcIndex).flags.ExpCount - ExpaDar

            End If
    
122         If ExpMult > 0 Then ExpaDar = ExpaDar * ExpMult

130         If ExpaDar > 0 Then
132             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
134                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

136                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

138                 Call WriteUpdateExp(UserIndex)
140                 Call CheckUserLevel(UserIndex)

                End If
            
142             Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(ExpaDar), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, RGB(0, 169, 255))

            End If

        End If

        
        Exit Sub

CalcularDarExp_Err:
144     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExp", Erl)
146     Resume Next
        
End Sub

Private Sub CalcularDarExpGrupal(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
        
        On Error GoTo CalcularDarExpGrupal_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim ExpaDar                 As Long
        Dim BonificacionGrupo       As Single
        Dim CantidadMiembrosValidos As Integer
        Dim i                       As Long
        Dim index                   As Integer

        'If UserList(UserIndex).Grupo.EnGrupo Then
        '[Nacho] Chekeamos que las variables sean validas para las operaciones
100     If NpcIndex = 0 Then Exit Sub
102     If UserIndex = 0 Then Exit Sub
104     If ElDaño <= 0 Then ElDaño = 0
106     If NpcList(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
108     If ElDaño > NpcList(NpcIndex).Stats.MinHp Then ElDaño = NpcList(NpcIndex).Stats.MinHp
    
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
110     ExpaDar = CLng((ElDaño) * (NpcList(NpcIndex).GiveEXP / NpcList(NpcIndex).Stats.MaxHp))

112     If ExpaDar <= 0 Then Exit Sub

        '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
114     If ExpaDar > NpcList(NpcIndex).flags.ExpCount Then
116         ExpaDar = NpcList(NpcIndex).flags.ExpCount
118         NpcList(NpcIndex).flags.ExpCount = 0
        Else
120         NpcList(NpcIndex).flags.ExpCount = NpcList(NpcIndex).flags.ExpCount - ExpaDar
        End If
        
122     For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
124         index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)
126         If UserList(index).flags.Muerto = 0 Then
128             If UserList(UserIndex).pos.Map = UserList(index).pos.Map Then
130                 If Abs(UserList(UserIndex).pos.x - UserList(index).pos.x) < 20 Then
132                     If Abs(UserList(UserIndex).pos.y - UserList(index).pos.y) < 20 Then
                            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then 'hay una var del lvl max?
134                             CantidadMiembrosValidos = CantidadMiembrosValidos + 1
                            End If
                        End If
                    End If
                End If
            End If
        Next
    
160     If ExpMult > 0 Then
162         ExpaDar = ExpaDar * ExpMult
        End If

166     ExpaDar = ExpaDar / CantidadMiembrosValidos

        Dim ExpUser As Long

168     If ExpaDar > 0 Then
170         For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
172             index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)
    
174             If UserList(index).flags.Muerto = 0 Then
176                 If Distancia(UserList(UserIndex).pos, UserList(index).pos) < 20 Then

186                     ExpUser = ExpaDar * ExpMult

188                     If UserList(index).Stats.ELV < STAT_MAXELV Then
190                         UserList(index).Stats.Exp = UserList(index).Stats.Exp + ExpUser

192                         If UserList(index).Stats.Exp > MAXEXP Then UserList(index).Stats.Exp = MAXEXP

194                         If UserList(index).ChatCombate = 1 Then
196                             Call WriteLocaleMsg(index, "141", FontTypeNames.FONTTYPE_EXP, ExpUser)

                            End If

198                         Call WriteUpdateExp(index)
200                         Call CheckUserLevel(index)

                        End If
    
                    Else
    
                        'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
202                     If UserList(index).ChatCombate = 1 Then
204                         Call WriteLocaleMsg(index, "69", FontTypeNames.FONTTYPE_New_GRUPO)
    
                        End If
    
                    End If
    
                Else
    
206                 If UserList(index).ChatCombate = 1 Then
208                     Call WriteConsoleMsg(index, "Estás muerto, no has ganado experencia del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
    
                    End If
    
                End If
    
210         Next i
        End If

        'Else
        '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, experencia perdida.", FontTypeNames.FONTTYPE_New_GRUPO)
        'End If

        
        Exit Sub

CalcularDarExpGrupal_Err:
212     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExpGrupal", Erl)
214     Resume Next
        
End Sub

Private Sub CalcularDarOroGrupal(ByVal UserIndex As Integer, ByVal GiveGold As Long)
        
        On Error GoTo CalcularDarOroGrupal_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/09/06 Nacho
        'Reescribi gran parte del Sub
        'Ahora, da toda la experiencia del npc mientras este vivo.
        '***************************************************
        Dim OroDar            As Long

124     OroDar = GiveGold * OroMult

        Dim orobackup As Long

130     orobackup = OroDar

        Dim i     As Byte

        Dim index As Byte

132     OroDar = OroDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

134     For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
136         index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)

138         If UserList(index).flags.Muerto = 0 Then
140             If UserList(UserIndex).pos.Map = UserList(index).pos.Map Then
142                 If OroDar > 0 Then

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
154     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CalcularDarOroGrupal", Erl)
156     Resume Next
        
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
        On Error GoTo ErrHandler

        Dim tOrg As eTrigger
        Dim tDst As eTrigger

100     tOrg = MapData(UserList(Origen).pos.Map, UserList(Origen).pos.x, UserList(Origen).pos.y).trigger
102     tDst = MapData(UserList(Destino).pos.Map, UserList(Destino).pos.x, UserList(Destino).pos.y).trigger
    
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
116     LogError ("Error en TriggerZonaPelea - " & Err.Description)

End Function

Private Sub UserDañoEspecial(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    On Error GoTo UserDañoEspecial_Err

    Dim ArmaObjInd As Integer, ObjInd As Integer
    ArmaObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    ObjInd = 0

    If ArmaObjInd = 0 Then
     ArmaObjInd = UserList(AtacanteIndex).Invent.NudilloObjIndex

    End If

    ' Preguntamos una vez mas, si no tiene Nudillos o Arma, no tiene sentido seguir.
    If ArmaObjInd = 0 Then
      Exit Sub
    End If

    If ObjData(ArmaObjInd).Proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
    End If

    Dim puedeEnvenenar, puedeEstupidizar, puedeIncinierar, puedeParalizar As Boolean
    puedeEnvenenar = (UserList(AtacanteIndex).flags.Envenena > 0) Or (ObjInd > 0 And ObjData(ObjInd).Envenena)
    puedeEstupidizar = (UserList(AtacanteIndex).flags.Estupidiza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Estupidiza)
    puedeIncinierar = (UserList(AtacanteIndex).flags.incinera > 0) Or (ObjInd > 0 And ObjData(ObjInd).incinera)
    puedeParalizar = (UserList(AtacanteIndex).flags.Paraliza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Paraliza)

    If puedeEnvenenar And (UserList(VictimaIndex).flags.Envenenado = 0) Then
        If RandomNumber(1, 100) < 30 Then
            UserList(VictimaIndex).flags.Envenenado = ObjData(ObjInd).Envenena
            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha envenenado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has envenenado a " & UserList(VictimaIndex).name & "!")
            
            Exit Sub
        End If
    End If

    If puedeIncinierar And (UserList(VictimaIndex).flags.Incinerado = 0) Then
        If RandomNumber(1, 100) < 10 Then
            UserList(VictimaIndex).flags.Incinerado = 1
            UserList(VictimaIndex).Counters.Incineracion = 1
            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha Incinerado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has Incinerado a " & UserList(VictimaIndex).name & "!")
            
            Exit Sub
        End If
    End If

    If puedeParalizar And (UserList(VictimaIndex).flags.Paralizado = 0) Then
        If RandomNumber(1, 100) < 10 Then
            UserList(VictimaIndex).flags.Paralizado = 1
            UserList(VictimaIndex).Counters.Paralisis = 6

            Call WriteParalizeOK(VictimaIndex)
            Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 8, 0))

            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha paralizado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has paralizado a " & UserList(VictimaIndex).name & "!")

            Exit Sub
        End If
    End If

    If puedeEstupidizar And (UserList(VictimaIndex).flags.Estupidez = 0) Then
        If RandomNumber(1, 100) < 13 Then
            UserList(VictimaIndex).flags.Estupidez = 1
            UserList(VictimaIndex).Counters.Estupidez = 3 ' segundos?

            Call WriteDumb(VictimaIndex)
            Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageParticleFX(UserList(VictimaIndex).Char.CharIndex, 30, 30, False))

            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha estupidizado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has estupidizado a " & UserList(VictimaIndex).name & "!")

            Exit Sub
        End If
    End If

    Exit Sub

UserDañoEspecial_Err:
    Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UserDañoEspecial", Erl)
    Resume Next

End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    'Reaccion de las mascotas
        
    On Error GoTo AllMascotasAtacanUser_Err

    Dim iCount As Long
    Dim mascotaIndex As Integer
    
    With UserList(Maestro)
    
        For iCount = 1 To MAXMASCOTAS
            mascotaIndex = .MascotasIndex(iCount)
            
            If mascotaIndex > 0 Then
                If NpcList(mascotaIndex).flags.AtacaUsuarios Then
                    NpcList(mascotaIndex).flags.AttackedBy = UserList(victim).name
                    NpcList(mascotaIndex).Target = victim
                    NpcList(mascotaIndex).Movement = TipoAI.NpcDefensa
                    NpcList(mascotaIndex).Hostile = 1
                End If
                    
            End If
            
        Next iCount
    
    End With
        
    Exit Sub

AllMascotasAtacanUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanUser", Erl)
        
End Sub

Public Sub AllMascotasAtacanNPC(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo AllMascotasAtacanNPC_Err
        
        Dim j As Long
        Dim mascotaIdx As Integer
        
        For j = 1 To MAXMASCOTAS
            mascotaIdx = UserList(UserIndex).MascotasIndex(j)
            
            If mascotaIdx > 0 And mascotaIdx <> NpcIndex Then
                With NpcList(mascotaIdx)
                    
                    If .flags.AtacaNPCs And .TargetNPC = 0 Then
                        .TargetNPC = NpcIndex
                        .Movement = TipoAI.NpcAtacaNpc
                    End If
            
                End With
            End If
        Next j
        
        Exit Sub

AllMascotasAtacanNPC_Err:
116     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanNPC", Erl)
        
End Sub

Private Function PuedeDesequiparDeUnGolpe(ByVal UserIndex As Integer) As Boolean
    On Error GoTo PuedeDesequiparDeUnGolpe_Err
    
    With UserList(UserIndex)
        Select Case .clase
    
        Case eClass.Bandit, eClass.Thief
            PuedeDesequiparDeUnGolpe = (.Stats.UserSkills(eSkill.Wrestling) >= 100) And (.Invent.WeaponEqpObjIndex = 0)

        Case Else
            PuedeDesequiparDeUnGolpe = False
    
        End Select
            
    End With
        
    Exit Function

PuedeDesequiparDeUnGolpe_Err:
    Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeDesequiparDeUnGolpe", Erl)
    Resume Next
        
End Function

Private Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo PuedeApuñalar_Err
        
        With UserList(UserIndex)

100         If .Invent.WeaponEqpObjIndex > 0 Then
102             PuedeApuñalar = (.clase = eClass.Assasin Or .Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) And ObjData(.Invent.WeaponEqpObjIndex).Apuñala = 1
            End If
            
        End With
        
        Exit Function

PuedeApuñalar_Err:
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeApuñalar", Erl)
108     Resume Next
        
End Function

Private Function PuedeGolpeCritico(ByVal UserIndex As Integer) As Boolean
        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo PuedeGolpeCritico_Err
        
        With UserList(UserIndex)
    
100         If .Invent.WeaponEqpObjIndex > 0 Then
                ' Esto me parece que esta MAL; subtipo 2 es incinera :/
102             PuedeGolpeCritico = .clase = eClass.Bandit And ObjData(.Invent.WeaponEqpObjIndex).Subtipo = 2
            End If
            
        End With
        
        Exit Function

PuedeGolpeCritico_Err:
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeGolpeCritico", Erl)
108     Resume Next
        
End Function

Private Function ProbabilidadApuñalar(ByVal UserIndex As Integer) As Integer

        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo ProbabilidadApuñalar_Err

        With UserList(UserIndex)

100         Dim Skill  As Integer
102         Skill = .Stats.UserSkills(eSkill.Apuñalar)
        
104         Select Case .clase
    
                Case eClass.Assasin '20%
106                 ProbabilidadApuñalar = 0.33 * Skill
    
108             Case eClass.Pirat, eClass.Hunter '15%
110                 ProbabilidadApuñalar = 0.15 * Skill
    
112             Case Else ' 10%
114                 ProbabilidadApuñalar = 0.1 * Skill
    
            End Select
            
            ' Daga especial da +5 de prob. de apu
116         If ObjData(.Invent.WeaponEqpObjIndex).Subtipo = 42 Then
118             ProbabilidadApuñalar = ProbabilidadApuñalar + 5
            End If
            
        End With
        
        Exit Function

ProbabilidadApuñalar_Err:
120     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadApuñalar", Erl)
122     Resume Next
        
End Function

Private Function ProbabilidadGolpeCritico(ByVal UserIndex As Integer) As Integer
        On Error GoTo ProbabilidadGolpeCritico_Err

        ProbabilidadGolpeCritico = 0.2 * UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

        Exit Function

ProbabilidadGolpeCritico_Err:
132     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadGolpeCritico", Erl)
134     Resume Next

End Function

Private Function ProbabilidadDesequipar(ByVal UserIndex As Integer) As Integer
    On Error GoTo ProbabilidadDesequipar_Err

    With UserList(UserIndex)

        Select Case .clase
    
        Case eClass.Bandit
            ProbabilidadDesequipar = 0.2 * 100
    
        Case eClass.Thief
            ProbabilidadDesequipar = 0.33 * 100
    
        Case Else
            ProbabilidadDesequipar = 0
    
        End Select
               
    End With
        
    Exit Function

ProbabilidadDesequipar_Err:
    Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadDesequipar", Erl)
    Resume Next
        
End Function


' Helper function to simplify the code. Keep private!
Private Sub WriteCombatConsoleMsg(ByVal UserIndex As Integer, ByVal message As String)
        On Error GoTo WriteCombatConsoleMsg_Err

        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteConsoleMsg(UserIndex, message, FontTypeNames.FONTTYPE_FIGHT)
        End If

        Exit Sub

WriteCombatConsoleMsg_Err:
        Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.WriteCombatConsoleMsg", Erl)
        Resume Next

End Sub
