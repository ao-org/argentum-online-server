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
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModificadorEvasion", Erl)
104     Resume Next
        
End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueArmas_Err
        

100     ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

        
        Exit Function

ModificadorPoderAtaqueArmas_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueArmas", Erl)
104     Resume Next
        
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModificadorPoderAtaqueProyectiles_Err
        
    
100     ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

        
        Exit Function

ModificadorPoderAtaqueProyectiles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueProyectiles", Erl)
104     Resume Next
        
End Function

Function ModicadorDañoClaseArmas(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseArmas_Err
        
    
100     ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

        
        Exit Function

ModicadorDañoClaseArmas_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseArmas", Erl)
104     Resume Next
        
End Function
Function ModicadorApuñalarClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorApuñalarClase_Err
        
    
100     ModicadorApuñalarClase = ModClase(clase).ModApuñalar

        
        Exit Function

ModicadorApuñalarClase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorApuñalarClase", Erl)
104     Resume Next
        
End Function

Function ModicadorDañoClaseWrestling(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseWrestling_Err
        
        
100     ModicadorDañoClaseWrestling = ModClase(clase).DañoWrestling

        
        Exit Function

ModicadorDañoClaseWrestling_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseWrestling", Erl)
104     Resume Next
        
End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As eClass) As Single
        
        On Error GoTo ModicadorDañoClaseProyectiles_Err
        
        
100     ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

        
        Exit Function

ModicadorDañoClaseProyectiles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseProyectiles", Erl)
104     Resume Next
        
End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single
        
        On Error GoTo ModEvasionDeEscudoClase_Err
        

100     ModEvasionDeEscudoClase = ModClase(clase).Escudo

        
        Exit Function

ModEvasionDeEscudoClase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ModEvasionDeEscudoClase", Erl)
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

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
        
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

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
        
        On Error GoTo PoderEvasionEscudo_Err
        

100     PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

        
        Exit Function

PoderEvasionEscudo_Err:
102     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderEvasionEscudo", Erl)
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
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderEvasion", Erl)
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
116     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueArma", Erl)
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
116     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueProyectil", Erl)
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
            SubirSkillDeArmaActual(UserIndex)
        End If

        Exit Function

UserImpactoNpc_Err:
130     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UserImpactoNpc", Erl)
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
136     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcImpacto", Erl)
138     Resume Next
        
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer) As Long

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
                ModifClase = ModicadorDañoClaseWrestling(.clase)
            
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
            If .clase = eClass.Pirat And .flags.Navegando Then
                CalcularDaño = CalcularDaño * 1.15
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
122             If RandomNumber(1, 100) <= ProbabilidadGolpeCritico(UserIndex) * 2 Then
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
136             If RandomNumber(1, 100) <= ProbabilidadApuñalar(UserIndex) * 2 Then
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

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
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
        
            'Si lo mato un guardia
174         If Status(UserIndex) = 2 And NpcList(NpcIndex).NPCtype = eNPCType.GuardiaReal Then

                ' Call RestarCriminalidad(UserIndex)
176             If Status(UserIndex) < 2 And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)

            End If
            
178         If NpcList(NpcIndex).MaestroUser > 0 Then
180             Call AllFollowAmo(NpcList(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
182             If NpcList(NpcIndex).Stats.Alineacion = 0 Then
184                 NpcList(NpcIndex).Movement = NpcList(NpcIndex).flags.OldMovement
186                 NpcList(NpcIndex).Hostile = NpcList(NpcIndex).flags.OldHostil
188                 NpcList(NpcIndex).flags.AttackedBy = vbNullString
                End If
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

116     If NpcList(NpcIndex).Target = 0 Then NpcList(NpcIndex).Target = UserIndex
    
118     If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex

120     NpcList(NpcIndex).CanAttack = 0
    
122     If NpcList(NpcIndex).flags.Snd1 > 0 Then
124         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd1, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
        End If
        
126     Call CancelExit(UserIndex)

128     If NpcImpacto(NpcIndex, UserIndex) Then
    
130         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        
132         If UserList(UserIndex).flags.Navegando = 0 Or UserList(UserIndex).flags.Montado = 0 Then
134             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))

            End If
        
136         Call NpcDaño(NpcIndex, UserIndex)

            '¿Puede envenenar?
138         If NpcList(NpcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, NpcList(NpcIndex).Veneno)
        Else
140         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharSwing(NpcList(NpcIndex).Char.CharIndex, False))

        End If

        '-----Tal vez suba los skills------
142     Call SubirSkill(UserIndex, Tacticas)
    
        'Controla el nivel del usuario
144     Call CheckUserLevel(UserIndex)
        

        Exit Function

NpcAtacaUser_Err:
146     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.NpcAtacaUser", Erl)
148     Resume Next
        
End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
        
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

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
        
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

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
        
        On Error GoTo NpcAtacaNpc_Err
        
 
        ' El npc puede atacar ???
100     If IntervaloPermiteAtacarNPC(Atacante) Then
102         NpcList(Atacante).CanAttack = 0

104         If cambiarMOvimiento Then
106             NpcList(Victima).TargetNPC = Atacante
108             NpcList(Victima).Movement = TipoAI.NpcAtacaNpc

            End If

        Else
            Exit Sub

        End If

110     If NpcList(Atacante).flags.Snd1 > 0 Then
112         Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(NpcList(Atacante).flags.Snd1, NpcList(Atacante).Pos.X, NpcList(Atacante).Pos.Y))

        End If

114     If NpcImpactoNpc(Atacante, Victima) Then
    
116         If NpcList(Victima).flags.Snd2 > 0 Then
118             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(NpcList(Victima).flags.Snd2, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))
            Else
120             Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))

            End If

122         Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, NpcList(Victima).Pos.X, NpcList(Victima).Pos.Y))
    
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
        

100     If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
            Exit Sub

        End If
    
102     If UserList(UserIndex).flags.invisible = 0 Then
104         Call NPCAtacado(NpcIndex, UserIndex)
        End If

106     If UserImpactoNpc(UserIndex, NpcIndex) Then
        
108         If NpcList(NpcIndex).flags.Snd2 > 0 Then
110             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))
            Else
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y))

            End If

114         If UserList(UserIndex).flags.Paraliza = 1 And NpcList(NpcIndex).flags.Paralizado = 0 Then

                Dim Probabilidad As Byte
    
116             Probabilidad = RandomNumber(1, 4)

118             If Probabilidad = 1 Then
120                 If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
122                     NpcList(NpcIndex).flags.Paralizado = 1
                        
124                     NpcList(NpcIndex).Contadores.Paralisis = IntervaloParalizado / 3

126                     If UserList(UserIndex).ChatCombate = 1 Then
                            'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
128                         Call WriteLocaleMsg(UserIndex, "136", FontTypeNames.FONTTYPE_FIGHT)

                        End If

130                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.CharIndex, 8, 0))
                                     
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
            
            Dim sendto As SendTarget

138         If UserList(UserIndex).clase = eClass.Hunter And UserList(UserIndex).flags.Oculto = 0 Then
140             sendto = SendTarget.ToPCArea
            Else
142             sendto = SendTarget.ToIndex
            End If

144         Call SendData(sendto, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))

        End If

        
        Exit Sub

UsuarioAtacaNpc_Err:
146     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaNpc", Erl)
148     Resume Next
        
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
            
            Dim sendto As SendTarget

112         If UserList(UserIndex).clase = eClass.Hunter And UserList(UserIndex).flags.Oculto = 0 Then
114             sendto = SendTarget.ToPCArea
            Else
116             sendto = SendTarget.ToIndex
            End If
            
118         Call SendData(sendto, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))
            
120         UsuarioAtacaNpcFunction = 2

        End If

        
        Exit Function

UsuarioAtacaNpcFunction_Err:
122     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaNpcFunction", Erl)
124     Resume Next
        
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

            'Look for NPC
142         ElseIf MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex > 0 Then

144             index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex

146             If NpcList(index).Attackable Then
148                 If NpcList(index).MaestroUser > 0 And MapInfo(NpcList(index).Pos.Map).Seguro = 1 Then
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
          SubirSkillDeArmaActual(AtacanteIndex)

        Else ' Falló
            If RandomNumber(1, 100) <= ProbRechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
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

102     If Distancia(UserList(AtacanteIndex).Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
104         Call WriteLocaleMsg(AtacanteIndex, "8", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(atacanteindex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If

108     Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)

110     If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then

114         Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).Pos.X, UserList(AtacanteIndex).Pos.Y))

116         If UserList(VictimaIndex).flags.Navegando = 0 Or UserList(VictimaIndex).flags.Montado = 0 Then
118             Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If

            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en accion
120         If UserList(AtacanteIndex).clase = eClass.Bandit Then
122             Call DoDesequipar(AtacanteIndex, VictimaIndex)

                'y ahora, el ladron puede llegar a paralizar con el golpe.
124         ElseIf UserList(AtacanteIndex).clase = eClass.Thief Then
126             Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If

128         Call UserDañoUser(AtacanteIndex, VictimaIndex)

        Else

130         If UserList(AtacanteIndex).clase = eClass.Hunter And UserList(AtacanteIndex).flags.Oculto = 0 Then
132             sendto = SendTarget.ToPCArea
            Else
134             sendto = SendTarget.ToIndex
            End If

136         Call SendData(sendto, AtacanteIndex, PrepareMessageCharSwing(UserList(AtacanteIndex).Char.CharIndex))

        End If

        Exit Sub

UsuarioAtacaUsuario_Err:
138     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaUsuario", Erl)
140     Resume Next

End Sub

Private Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

        ' Reescrito por WyroX - 16/01/2021
        
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

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
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


130     Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
132     Call AllMascotasAtacanUser(VictimIndex, attackerIndex)

        Exit Sub

UsuarioAtacadoPorUsuario_Err:
136     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacadoPorUsuario", Erl)
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
        
        ' No podes atacar si estas en consulta
112     If UserList(attackerIndex).flags.EnConsulta Then
114         Call WriteConsoleMsg(attackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
    
        End If
        
        ' No podes atacar si esta en consulta
116     If UserList(VictimIndex).flags.EnConsulta Then
118         Call WriteConsoleMsg(attackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
    
        End If
        
120     If UserList(attackerIndex).flags.Maldicion = 1 Then
122         Call WriteConsoleMsg(attackerIndex, "¡Estas maldito! No podes atacar.", FontTypeNames.FONTTYPE_INFOIAO)
124         PuedeAtacar = False
            Exit Function

        End If

        'Es miembro del grupo?
126     'If UserList(attackerIndex).Grupo.EnGrupo = True Then

        '    Dim i As Byte

128     '    For i = 1 To UserList(UserList(attackerIndex).Grupo.Lider).Grupo.CantidadMiembros
    
130     '        If UserList(UserList(attackerIndex).Grupo.Lider).Grupo.Miembros(i) = VictimIndex Then
132     '            PuedeAtacar = False
134     '            Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        '            Exit Function

        '        End If

136     '    Next i

        'End If

        'Estamos en una Arena? o un trigger zona segura?
138     T = TriggerZonaPelea(attackerIndex, VictimIndex)

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

        'Consejeros no pueden atacar
        'If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
        '    PuedeAtacar = False
        '    Exit Sub
        'End If

150     If UserList(attackerIndex).GuildIndex <> 0 Then
152         If UserList(attackerIndex).flags.SeguroClan Then
154             If UserList(attackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
156                 Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu clan. Para hacerlo debes desactivar el seguro de clan.", FontTypeNames.FONTTYPE_INFOIAO)
158                 PuedeAtacar = False
                    Exit Function

                End If

            End If

        End If
        
        'Solo administradores pueden atacar a usuarios (PARA TESTING)
160     If (UserList(attackerIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Admin)) = 0 Then
162         PuedeAtacar = False
            Exit Function
        End If
        
        'Estas queriendo atacar a un GM?
164     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

166     If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
168         If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
170         PuedeAtacar = False
            Exit Function

        End If

        'Sos un Armada atacando un ciudadano?
172     If (Status(VictimIndex) = 1) And (esArmada(attackerIndex)) Then
174         Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
176         PuedeAtacar = False
            Exit Function

        End If

        'Tenes puesto el seguro?
178     If UserList(attackerIndex).flags.Seguro Then
180         If Status(VictimIndex) = 1 Then
182             Call WriteConsoleMsg(attackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
184             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Es un ciuda queriando atacar un imperial?
186     If UserList(attackerIndex).flags.Seguro Then
188         If (Status(attackerIndex) = 1) And (esArmada(VictimIndex)) Then
190             Call WriteConsoleMsg(attackerIndex, "Los ciudadanos no pueden atacar a los soldados imperiales.", FontTypeNames.FONTTYPE_WARNING)
192             PuedeAtacar = False
                Exit Function

            End If

        End If

        'Estas en un Mapa Seguro?
194     If MapInfo(UserList(VictimIndex).Pos.Map).Seguro = 1 Then

196         If esArmada(attackerIndex) Then
198             If UserList(attackerIndex).Faccion.RecompensasReal >= 3 Then
200                 If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
202                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
204                     PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

206         If esCaos(attackerIndex) Then
208             If UserList(attackerIndex).Faccion.RecompensasCaos >= 3 Then
210                 If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
212                     Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
214                     PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                        Exit Function

                    End If

                End If

            End If

216         Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
218         PuedeAtacar = False
            Exit Function

        End If

        'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
220     If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(attackerIndex).Pos.Map, UserList(attackerIndex).Pos.X, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
222         Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
224         PuedeAtacar = False
            Exit Function

        End If

226     PuedeAtacar = True

        
        Exit Function

PuedeAtacar_Err:
228     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeAtacar", Erl)
230     Resume Next
        
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
106     If (UserList(attackerIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) = 0 Then
108         PuedeAtacarNPC = False
            Exit Function
        End If
        
        ' No podes atacar si estas en consulta
110     If UserList(attackerIndex).flags.EnConsulta Then
112         Call WriteConsoleMsg(attackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarNPC = False
            Exit Function

        End If
        
        'Es una criatura atacable?
114     If NpcList(NpcIndex).Attackable = 0 Then
            'No es una criatura atacable
116         Call WriteConsoleMsg(attackerIndex, "No podés atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
118         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es valida la distancia a la cual estamos atacando?
120     If Distancia(UserList(attackerIndex).Pos, NpcList(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
122         Call WriteLocaleMsg(attackerIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
124         PuedeAtacarNPC = False
            Exit Function

        End If

        'Es una criatura No-Hostil?
126     If NpcList(NpcIndex).Hostile = 0 Then
            'Es una criatura No-Hostil.
            'Es Guardia del Caos?

128         If NpcList(NpcIndex).NPCtype = eNPCType.Guardiascaos Then

                'Lo quiere atacar un caos?
130             If esCaos(attackerIndex) Then
132                 Call WriteConsoleMsg(attackerIndex, "No podés atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
134                 PuedeAtacarNPC = False
                    Exit Function

                End If

136             If Status(attackerIndex) = 1 Then
138                 PuedeAtacarNPC = True
                    Exit Function

                End If
        
            End If

            'Es guardia Real?
140         If NpcList(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                'Lo quiere atacar un Armada?
        
142             If esCaos(attackerIndex) Then
144                 PuedeAtacarNPC = True
                    Exit Function

                End If
        
146             If esArmada(attackerIndex) Then
148                 Call WriteConsoleMsg(attackerIndex, "No podés atacar Guardias Reales siendo Armada Real", FontTypeNames.FONTTYPE_INFO)
150                 PuedeAtacarNPC = False
                    Exit Function

                End If
        
                'Tienes el seguro puesto?
152             If UserList(attackerIndex).flags.Seguro And Status(attackerIndex) = 1 Then
154                 Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para poder Atacar Guardias Reales utilizando /seg", FontTypeNames.FONTTYPE_INFO)
156                 PuedeAtacarNPC = False
                    Exit Function
                Else
158                 Call WriteConsoleMsg(attackerIndex, "Atacaste un Guardia Real! Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
160                 Call VolverCriminal(attackerIndex)
162                 PuedeAtacarNPC = True
                    Exit Function

                End If

            End If

        End If
        
        'Es el NPC mascota de alguien?
164     If NpcList(NpcIndex).MaestroUser > 0 Then
166         If UserList(NpcList(NpcIndex).MaestroUser).Faccion.Status = 1 Then
                'Es mascota de un Ciudadano.
168             If UserList(attackerIndex).Faccion.Status = 1 Then
                    'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
170                 If UserList(attackerIndex).flags.Seguro Then
                        'El atacante tiene el seguro puesto. No puede atacar.
172                     Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro de combate.", FontTypeNames.FONTTYPE_INFO)
174                     PuedeAtacarNPC = False
                        Exit Function
                    Else
                        'El atacante no tiene el seguro puesto. Recibe penalización.
176                     Call WriteConsoleMsg(attackerIndex, "Has atacado la mascota de un ciudadano. Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
178                     Call VolverCriminal(attackerIndex)
180                     PuedeAtacarNPC = True
                        Exit Function
                    End If
                Else
                    'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
182                 If UserList(attackerIndex).flags.Seguro Then
184                     Call WriteConsoleMsg(attackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro de combate.", FontTypeNames.FONTTYPE_INFO)
186                     PuedeAtacarNPC = False
                        Exit Function
                    End If
                End If
            Else
                'Es mascota de un Criminal.
188             If esCaos(NpcList(NpcIndex).MaestroUser) Then
                    'Es Caos el Dueño.
190                 If esCaos(attackerIndex) Then
                        'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
192                     Call WriteConsoleMsg(attackerIndex, "Los miembros de la Legión Oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
194                     PuedeAtacarNPC = False
                        Exit Function
                    End If
                End If
            End If
        End If
        
        'Es el Rey Preatoriano?
196     If NpcList(NpcIndex).NPCtype = eNPCType.Pretoriano Then
198         If Not ClanPretoriano(NpcList(NpcIndex).ClanIndex).CanAtackMember(NpcIndex) Then
200             Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
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

            Dim ExpaDar As Long
    
            '[Nacho] Chekeamos que las variables sean validas para las operaciones
104         If ElDaño <= 0 Then ElDaño = 0
106         If NpcList(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub

            '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
110         ExpaDar = ElDaño * NpcList(NpcIndex).GiveEXP / NpcList(NpcIndex).Stats.MaxHp

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
    
122         If ExpMult > 0 Then
124             ExpaDar = ExpaDar * ExpMult * UserList(UserIndex).flags.ScrollExp
    
            End If
    
126         If UserList(UserIndex).donador.activo = 1 Then
128             ExpaDar = ExpaDar * 1.1

            End If

130         If ExpaDar > 0 Then
132             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
134                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

136                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

138                 Call WriteUpdateExp(UserIndex)
140                 Call CheckUserLevel(UserIndex)

                End If
            
142             Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(ExpaDar), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, RGB(0, 169, 255))

            End If

        End If

        
        Exit Sub

CalcularDarExp_Err:
144     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExp", Erl)
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
128             If UserList(UserIndex).Pos.Map = UserList(index).Pos.Map Then
130                 If Abs(UserList(UserIndex).Pos.X - UserList(index).Pos.X) < 20 Then
132                     If Abs(UserList(UserIndex).Pos.Y - UserList(index).Pos.Y) < 20 Then
134                         CantidadMiembrosValidos = CantidadMiembrosValidos + 1
                        End If
                    End If
                End If
            End If
        Next
    
136     Select Case CantidadMiembrosValidos
    
            Case 1
138             BonificacionGrupo = 1

140         Case 2
142             BonificacionGrupo = 1.2

144         Case 3
146             BonificacionGrupo = 1.4

148         Case 4
150             BonificacionGrupo = 1.6

152         Case 5
154             BonificacionGrupo = 1.8

156         Case Else
158             BonificacionGrupo = 2
                
        End Select
 
160     If ExpMult > 0 Then
162         ExpaDar = ExpaDar * ExpMult
        End If

164     ExpaDar = ExpaDar * BonificacionGrupo

166     ExpaDar = ExpaDar / CantidadMiembrosValidos
    
        Dim ExpUser As Long
    
168     If ExpaDar > 0 Then
170         For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
172             index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)
    
174             If UserList(index).flags.Muerto = 0 Then
176                 If Distancia(UserList(UserIndex).Pos, UserList(index).Pos) < 20 Then

178                     ExpUser = 0

180                     If UserList(index).donador.activo = 1 Then
182                         ExpUser = ExpaDar * 1.1
                        Else
184                         ExpUser = ExpaDar
                        End If
                    
186                     ExpUser = ExpUser * UserList(index).flags.ScrollExp
                
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
154     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CalcularDarOroGrupal", Erl)
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
116     LogError ("Error en TriggerZonaPelea - " & Err.Description)

End Function

Private Sub UserDañoEspecial(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
        On Error GoTo UserDañoEspecial_Err

        Dim ArmaObjInd As Integer, ObjInd As Integer
        Dim HuboEfecto As Boolean

        HuboEfecto = False
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
        puedeEnvenenar   = (UserList(AtacanteIndex).flags.Envenena > 0)   Or (ObjInd > 0 And ObjData(ObjInd).Envenena)
        puedeEstupidizar = (UserList(AtacanteIndex).flags.Estupidiza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Estupidiza)
        puedeIncinierar  = (UserList(AtacanteIndex).flags.incinera > 0)   Or (ObjInd > 0 And ObjData(ObjInd).incinera)
        puedeParalizar   = (UserList(AtacanteIndex).flags.Paraliza > 0)   Or (ObjInd > 0 And ObjData(ObjInd).Paraliza)

        If puedeEnvenenar And (UserList(VictimaIndex).flags.Envenenado = 0) And Not HuboEfecto Then
            If RandomNumber(1, 100) < 30 Then
                UserList(VictimaIndex).flags.Envenenado = ObjData(ObjInd).Envenena
                Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha envenenado!")
                Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has envenenado a " & UserList(VictimaIndex).name & "!")
                HuboEfecto = True

            End If
        End If

        If puedeIncinierar And (UserList(VictimaIndex).flags.Incinerado = 0) And Not HuboEfecto Then
            If RandomNumber(1, 100) < 10 Then
                UserList(VictimaIndex).flags.Incinerado = 1
                UserList(VictimaIndex).Counters.Incineracion = 1
                Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha Incinerado!")
                Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has Incinerado a " & UserList(VictimaIndex).name & "!")
                HuboEfecto = True

            End If
        End If

        If puedeParalizar And (UserList(VictimaIndex).flags.Paralizado = 0) And Not HuboEfecto Then
            If RandomNumber(1, 100) < 10 Then
                UserList(VictimaIndex).flags.Paralizado = 1
                UserList(VictimaIndex).Counters.Paralisis = 6

                Call WriteParalizeOK(VictimaIndex)
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, 8, 0))

                Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha paralizado!")
                Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has paralizado a " & UserList(VictimaIndex).name & "!")

                HuboEfecto = True

            End If
        End If

        If puedeEstupidizar And (UserList(VictimaIndex).flags.Estupidez = 0) And Not HuboEfecto Then
            If RandomNumber(1, 100) < 8 Then
                UserList(VictimaIndex).flags.Estupidez = 1
                UserList(VictimaIndex).Counters.Estupidez = 5

                Call WriteDumb(VictimaIndex)
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageParticleFX(UserList(VictimaIndex).Char.CharIndex, 30, 30, False))

                Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha estupidizado!")
                Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has estupidizado a " & UserList(VictimaIndex).name & "!")

                HuboEfecto = True
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
    
        
        Dim iCount As Integer
    
100     For iCount = 1 To MAXMASCOTAS
102         If UserList(Maestro).MascotasIndex(iCount) > 0 Then
104             NpcList(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).name
106             NpcList(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
108             NpcList(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
            End If
110     Next iCount
        
        Exit Sub

AllMascotasAtacanUser_Err:
112     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanUser", Erl)

        
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
        
        On Error GoTo CheckPets_Err
    
        
        Dim j As Integer
    
100     For j = 1 To MAXMASCOTAS
102         If UserList(UserIndex).MascotasIndex(j) > 0 Then
104            If UserList(UserIndex).MascotasIndex(j) <> NpcIndex Then
106             If CheckElementales Or (NpcList(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And NpcList(UserList(UserIndex).MascotasIndex(j)).Numero <> ELEMENTALVIENTO) Then
108                 If NpcList(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0 Then NpcList(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex
110                 NpcList(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                End If
               End If
            End If
112     Next j
        
        Exit Sub

CheckPets_Err:
114     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.CheckPets", Erl)

        
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
108     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.AllFollowAmo", Erl)

        
End Sub

Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
        
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

Function PuedeGolpeCritico(ByVal UserIndex As Integer) As Boolean
        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo PuedeGolpeCritico_Err
        
        With UserList(UserIndex)

100         If .Invent.WeaponEqpObjIndex > 0 Then
102             PuedeGolpeCritico = .clase = eClass.Bandit And ObjData(.Invent.WeaponEqpObjIndex).Subtipo = 2
            End If
            
        End With
        
        Exit Function

PuedeGolpeCritico_Err:
106     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.PuedeGolpeCritico", Erl)
108     Resume Next
        
End Function

Public Function ProbabilidadApuñalar(ByVal UserIndex As Integer) As Integer

        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo ProbabilidadApuñalar_Err

        With UserList(UserIndex)

100         Dim Skill  As Integer
102         Skill = .Stats.UserSkills(eSkill.Apuñalar)
        
104         Select Case .clase
    
                Case eClass.Assasin '20%
106                 ProbabilidadApuñalar = 0.2 * Skill
    
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

Public Function ProbabilidadGolpeCritico(ByVal UserIndex As Integer) As Integer
        ' Autor: WyroX - 16/01/2021
        
        On Error GoTo ProbabilidadGolpeCritico_Err

100     With UserList(UserIndex)

102         ProbabilidadGolpeCritico = 0.2 * .Stats.UserSkills(eSkill.Wrestling)

        End With

        Exit Function

ProbabilidadGolpeCritico_Err:
132     Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadGolpeCritico", Erl)

134     Resume Next
        
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
