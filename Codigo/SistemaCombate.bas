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

    ModificadorEvasion = ModClase(clase).Evasion

End Function

Function ModificadorPoderAtaqueArmas(ByVal clase As eClass) As Single

    ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas

End Function

Function ModificadorPoderAtaqueProyectiles(ByVal clase As eClass) As Single
    
    ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles

End Function

Function ModicadorDañoClaseArmas(ByVal clase As eClass) As Single
    
    ModicadorDañoClaseArmas = ModClase(clase).DañoArmas

End Function

Function ModicadorDañoClaseWrestling(ByVal clase As eClass) As Single
        
    ModicadorDañoClaseWrestling = ModClase(clase).DañoWrestling

End Function

Function ModicadorDañoClaseProyectiles(ByVal clase As eClass) As Single
        
    ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles

End Function

Function ModEvasionDeEscudoClase(ByVal clase As eClass) As Single

    ModEvasionDeEscudoClase = ModClase(clase).Escudo

End Function

Function Minimo(ByVal a As Single, ByVal b As Single) As Single

    If a > b Then
        Minimo = b
        Else:
        Minimo = a

    End If

End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MinimoInt = b
        Else:
        MinimoInt = a

    End If

End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single

    If a > b Then
        Maximo = a
        Else:
        Maximo = b

    End If

End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer

    If a > b Then
        MaximoInt = a
        Else:
        MaximoInt = b

    End If

End Function

Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2

End Function

Function PoderEvasion(ByVal UserIndex As Integer) As Long

    Dim lTemp As Long

    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorEvasion(.clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))

    End With

End Function

Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long

    Dim PoderAtaqueTemp As Long

    If UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 31 Then
        PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Armas) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Armas) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    Else
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Armas) + (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))

    End If

    PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function

Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long

    Dim PoderAtaqueTemp As Long

    If UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
        PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
    Else
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Proyectiles) + (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))

    End If

    PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function

Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long

    Dim PoderAtaqueTemp As Long

    If UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 31 Then
        PoderAtaqueTemp = (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 61 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad)) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) < 91 Then
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + (2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    Else
        PoderAtaqueTemp = ((UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) + (3 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))) * ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))

    End If

    PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(UserIndex).Stats.ELV) - 12, 0)))

End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean

    Dim PoderAtaque As Long

    Dim Arma        As Integer

    Dim proyectil   As Boolean

    Dim ProbExito   As Long

    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex

    If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

    If Arma > 0 Then 'Usando un arma
        If proyectil Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)

        End If

    Else 'Peleando con puños
        PoderAtaque = PoderAtaqueWrestling(UserIndex)

    End If

    ProbExito = Maximo(10, Minimo(90, 70 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.1)))

    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

    If UserImpactoNpc Then
        If Arma <> 0 Then
            If proyectil Then
                Call SubirSkill(UserIndex, Proyectiles)
            Else
                Call SubirSkill(UserIndex, Armas)

            End If

        Else
            Call SubirSkill(UserIndex, Wrestling)

        End If

    End If

End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

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

    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)

    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)

    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

    ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.2)))

    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If Not NpcImpacto Then
            If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)

                If Rechazo = True Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))

                    If UserList(UserIndex).ChatCombate = 1 Then
                        Call WriteBlockedWithShieldUser(UserIndex)

                    End If

                    Call SubirSkill(UserIndex, Defensa)

                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 88, 0))
                End If

            End If

        End If

    End If

End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long

    Dim DañoArma As Long, DañoUsuario As Long, Arma As ObjData, ModifClase As Single

    Dim proyectil As ObjData

    Dim DañoMaxArma As Long

    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean

    matoDragon = False

    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        Arma = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex)
    
        ' Ataca a un npc?
        If NpcIndex > 0 Then

            'Usa la mata Dragones?
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            
                If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit
                    matoDragon = False ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                Else ' Sino es Dragon daño es 1
                    DañoArma = 1
                    DañoMaxArma = 1

                End If

            Else ' daño comun

                If Arma.proyectil = 1 Then
                    ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit

                    If Arma.Municion = 1 Then
                        proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma * 1.35
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)
                        DañoMaxArma = Arma.MaxHit
                        DañoMaxArma = DañoMaxArma * 1.35

                    End If

                Else
                    ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DañoArma = DañoArma * 1.35
                    DañoMaxArma = Arma.MaxHit
                    DañoMaxArma = DañoMaxArma * 1.35

                End If

            End If
    
        Else ' Ataca usuario

            If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                DañoArma = 1 ' Si usa la espada mataDragones daño es 1
                DañoMaxArma = 1
            Else

                If Arma.proyectil = 1 Then
                    ModifClase = ModicadorDañoClaseProyectiles(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit
                
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHit)
                        DañoMaxArma = Arma.MaxHit

                    End If

                Else
                    ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
                    DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit

                End If

            End If

        End If

    Else

        'Pablo (ToxicWaste)
        If UserList(UserIndex).Invent.NudilloSlot = 0 Then
            ModifClase = ModicadorDañoClaseWrestling(UserList(UserIndex).clase)
            DañoArma = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHit)
            DañoMaxArma = UserList(UserIndex).Stats.MaxHit
        Else
    
            ModifClase = ModicadorDañoClaseArmas(UserList(UserIndex).clase)
            Arma = ObjData(UserList(UserIndex).Invent.NudilloObjIndex)
            DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHit)
            DañoMaxArma = Arma.MaxHit

        End If

    End If

    If UserList(UserIndex).Invent.MagicoObjIndex = 707 And NpcIndex = 0 Then
        DañoUsuario = RandomNumber((UserList(UserIndex).Stats.MinHIT - ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CuantoAumento), (UserList(UserIndex).Stats.MaxHit - ObjData(UserList(UserIndex).Invent.MagicoObjIndex).CuantoAumento))
    Else
        DañoUsuario = RandomNumber(UserList(UserIndex).Stats.MinHIT, UserList(UserIndex).Stats.MaxHit)

    End If

    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    If matoDragon Then
        CalcularDaño = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
    Else
        CalcularDaño = ((3 * DañoArma) + ((DañoMaxArma / 5) * Maximo(0, (UserList(UserIndex).Stats.UserAtributos(Fuerza) - 15))) + DañoUsuario) * ModifClase
    
        'CalcularDaño = ((3 * 14) + ((14 / 5) * 20) + DañoUsuario) * ModifClase
        'CalcularDaño = (42 + (56 + 104) * ModifClase
        'CalcularDaño = 202 * 0.95  = 191      - defensas
    
        'CalcularDaño = 136
    End If

End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    Dim daño As Long

    Dim j As Integer

    Dim apudaño As Integer
    
    daño = CalcularDaño(UserIndex, NpcIndex)
    
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHit)

    If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then daño = daño + RandomNumber(ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MinHIT, ObjData(UserList(UserIndex).Invent.MonturaObjIndex).MaxHit)

    If PuedeApuñalar(UserIndex) Then
        Call SubirSkill(UserIndex, Apuñalar)
        apudaño = ApuñalarFunction(UserIndex, NpcIndex, 0, daño)

        ' daño = daño + apudaño
    End If
    
    daño = daño - Npclist(NpcIndex).Stats.def
    
    If daño < 0 Then daño = 0
    
    '[KEVIN]
    
    'If UserList(UserIndex).ChatCombate = 1 Then
    '    Call WriteUserHitNPC(UserIndex, daño)
    'End If
    
    If apudaño > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead("¡" & daño + apudaño & "!", Npclist(NpcIndex).Char.CharIndex, &HFFFF00))

        If UserList(UserIndex).ChatCombate = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & apudaño, FontTypeNames.FONTTYPE_FIGHT)
            
            Call WriteLocaleMsg(UserIndex, "212", FontTypeNames.FONTTYPE_FIGHT, apudaño)

        End If

    Else
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex))

    End If
    
    Call CalcularDarExp(UserIndex, NpcIndex, daño)
    Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
    '[/KEVIN]
     
    If Npclist(NpcIndex).Stats.MinHp <= 0 Then
            
        ' Si era un Dragon perdemos la espada mataDragones
        If Npclist(NpcIndex).NPCtype = DRAGON Then

            'Si tiene equipada la matadracos se la sacamos
            If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)

            End If

            ' If Npclist(NpcIndex).Stats.MaxHp > 100000 Then Call LogDesarrollo(UserList(UserIndex).name & " mató un dragón")
        End If
        
        Call MuereNpc(NpcIndex, UserIndex)

    End If

End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    Dim daño As Integer, Lugar As Integer, absorbido As Integer

    Dim antdaño As Integer, defbarco As Integer

    Dim obj As ObjData
    
    daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHit)
    antdaño = daño
    
    If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
        obj = ObjData(UserList(UserIndex).Invent.BarcoObjIndex)
        defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

    End If
    
    Dim defMontura As Integer

    If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
        obj = ObjData(UserList(UserIndex).Invent.MonturaObjIndex)
        defMontura = RandomNumber(obj.MinDef, obj.MaxDef)

    End If
    
    Lugar = RandomNumber(1, 6)
    
    Select Case Lugar

        Case PartesCuerpo.bCabeza

            'Si tiene casco absorbe el golpe
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                obj = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex)
                absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
                absorbido = absorbido + defbarco
                daño = daño - absorbido

                If daño < 1 Then daño = 1

            End If

        Case Else

            'Si tiene armadura absorbe el golpe
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then

                Dim Obj2 As ObjData

                obj = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex)

                If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
                    Obj2 = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex)
                    absorbido = RandomNumber(obj.MinDef + Obj2.MinDef, obj.MaxDef + Obj2.MaxDef)
                Else
                    absorbido = RandomNumber(obj.MinDef, obj.MaxDef)

                End If

                absorbido = absorbido + defbarco
                daño = daño - absorbido

                If daño < 1 Then daño = 1

            End If

    End Select
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectOverHead(daño, UserList(UserIndex).Char.CharIndex))

    If UserList(UserIndex).ChatCombate = 1 Then
        Call WriteNPCHitUser(UserIndex, Lugar, daño)

    End If

    If UserList(UserIndex).flags.Privilegios And PlayerType.user Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - daño
    
    If UserList(UserIndex).flags.Meditando Then
        If daño > Fix(UserList(UserIndex).Stats.MinHp / 100 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
            UserList(UserIndex).flags.Meditando = False
            Call WriteMeditateToggle(UserIndex)
            Call WriteLocaleMsg(UserIndex, "123", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.ParticulaFx, 0, True))
            UserList(UserIndex).Char.ParticulaFx = 0
            
        End If

    End If
    
    'Muere el usuario
    If UserList(UserIndex).Stats.MinHp <= 0 Then
    
        Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
        
        'Si lo mato un guardia
        If Status(UserIndex) = 2 And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then

            ' Call RestarCriminalidad(UserIndex)
            If Status(UserIndex) < 2 And UserList(UserIndex).Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(UserIndex)

        End If
        
        Call UserDie(UserIndex)
    
    End If

End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean

    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
    If (Not UserList(UserIndex).flags.Privilegios And PlayerType.user) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    
    ' El npc puede atacar ???
    
    If IntervaloPermiteAtacarNPC(NpcIndex) Then
    
        NpcAtacaUser = True

        If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    
        If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
    Else
        NpcAtacaUser = False
        Exit Function

    End If
    
    Npclist(NpcIndex).CanAttack = 0
    
    If Npclist(NpcIndex).flags.Snd1 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))

    End If
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        
        If UserList(UserIndex).flags.Navegando = 0 Or UserList(UserIndex).flags.Montado = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, FXSANGRE, 0))

        End If
        
        Call NpcDaño(NpcIndex, UserIndex)
        Call WriteUpdateHP(UserIndex)

        '¿Puede envenenar?
        If Npclist(NpcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, Npclist(NpcIndex).Veneno)
    Else
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharSwing(Npclist(NpcIndex).Char.CharIndex, False))

    End If

    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean

    Dim PoderAtt  As Long, PoderEva As Long

    Dim ProbExito As Long

    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtt - PoderEva) * 0.4)))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)

    Dim daño As Integer

    Dim ANpc As npc

    ANpc = Npclist(Atacante)
    
    daño = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHit)
    Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - daño
    
    If Npclist(Victima).Stats.MinHp < 1 Then
        
        If LenB(Npclist(Atacante).flags.AttackedBy) <> 0 Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement

        End If
        
    End If

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
 
    ' El npc puede atacar ???
    If IntervaloPermiteAtacarNPC(Atacante) Then
        Npclist(Atacante).CanAttack = 0

        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc

        End If

    Else
        Exit Sub

    End If

    If Npclist(Atacante).flags.Snd1 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1, Npclist(Atacante).Pos.x, Npclist(Atacante).Pos.Y))

    End If

    If NpcImpactoNpc(Atacante, Victima) Then
    
        If Npclist(Victima).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
        Else
            Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))

        End If

        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.x, Npclist(Victima).Pos.Y))
    
        Call NpcDañoNpc(Atacante, Victima)
    
    Else
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessageCharSwing(Npclist(Atacante).Char.CharIndex, False, True))

    End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        Exit Sub

    End If
    
    If UserList(UserIndex).flags.invisible = 0 Then
        Call NPCAtacado(NpcIndex, UserIndex)

    End If

    If UserImpactoNpc(UserIndex, NpcIndex) Then
        
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))

        End If

        If UserList(UserIndex).flags.Paraliza = 1 And Npclist(NpcIndex).flags.Paralizado = 0 Then

            Dim Probabilidad As Byte
    
            Probabilidad = RandomNumber(1, 4)

            If Probabilidad = 1 Then
                If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
                    Npclist(NpcIndex).flags.Paralizado = 1
                        
                    Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado / 3

                    If UserList(UserIndex).ChatCombate = 1 Then
                        'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
                        Call WriteLocaleMsg(UserIndex, "136", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.CharIndex, 8, 0))
                                     
                Else

                    If UserList(UserIndex).ChatCombate = 1 Then
                        'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If

        Call UserDañoNpc(UserIndex, NpcIndex)
       
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))

    End If

End Sub

Public Function UsuarioAtacaNpcFunction(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Byte

    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
        UsuarioAtacaNpcFunction = 0
        Exit Function

    End If
    
    Call NPCAtacado(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        Call UserDañoNpc(UserIndex, NpcIndex)
        UsuarioAtacaNpcFunction = 1
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex))
        UsuarioAtacaNpcFunction = 2

    End If

End Function

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Attack interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub

    'Quitamos stamina
    If UserList(UserIndex).Stats.MinSta < 10 Then
        'Call WriteConsoleMsg(UserIndex, "Estas muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
        Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    Call QuitarSta(UserIndex, RandomNumber(1, 10))
    
    If UserList(UserIndex).Counters.Trabajando Then
        Call WriteMacroTrabajoToggle(UserIndex, False)

    End If
        
    If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1

    'Movimiento de arma
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))
     
    'UserList(UserIndex).flags.PuedeAtacar = 0
    
    Dim AttackPos As WorldPos

    AttackPos = UserList(UserIndex).Pos
    Call HeadtoPos(UserList(UserIndex).Char.heading, AttackPos)
       
    'Exit if not legal
    If AttackPos.x >= XMinMapSize And AttackPos.x <= XMaxMapSize And AttackPos.Y >= YMinMapSize And AttackPos.Y <= YMaxMapSize Then

        Dim Index As Integer

        Index = MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).UserIndex
            
        'Look for user
        If Index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, Index)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(Index)

        'Look for NPC
        ElseIf MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex > 0 Then
            
            If Npclist(MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex).Attackable Then
                    
                Call UsuarioAtacaNpc(UserIndex, MapData(AttackPos.Map, AttackPos.x, AttackPos.Y).NpcIndex)
                Call WriteUpdateUserStats(UserIndex)
            Else
            
                Call WriteConsoleMsg(UserIndex, "No podés atacar a este NPC", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))
        End If

    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.CharIndex, True, False))

    End If

End Sub

Public Function UsuarioImpacto(ByVal atacanteindex As Integer, ByVal victimaindex As Integer) As Boolean
    
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
    
    If UserList(atacanteindex).flags.GolpeCertero = 1 Then
        UsuarioImpacto = True
        UserList(atacanteindex).flags.GolpeCertero = 0
        Exit Function

    End If
    
    SkillTacticas = UserList(victimaindex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(victimaindex).Stats.UserSkills(eSkill.Defensa)
    
    Arma = UserList(atacanteindex).Invent.WeaponEqpObjIndex

    If Arma > 0 Then
        proyectil = ObjData(Arma).proyectil = 1
    Else
        proyectil = False

    End If
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(victimaindex)
    
    If UserList(victimaindex).Invent.EscudoEqpObjIndex > 0 Then
        UserPoderEvasionEscudo = PoderEvasionEscudo(victimaindex)
        UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0

    End If
    
    'Esta usando un arma ???
    If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
        
        If proyectil Then
            PoderAtaque = PoderAtaqueProyectil(atacanteindex)
        Else
            PoderAtaque = PoderAtaqueArma(atacanteindex)

        End If

        ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))
       
    Else
        PoderAtaque = PoderAtaqueWrestling(atacanteindex)
        ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4)))
        
    End If

    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(victimaindex).Invent.EscudoEqpObjIndex > 0 Then
        
        'Fallo ???
        If UsuarioImpacto = False Then
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
          
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessagePlayWave(SND_ESCUDO, UserList(victimaindex).Pos.x, UserList(victimaindex).Pos.Y))
                Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEscudoMov(UserList(victimaindex).Char.CharIndex))

                If UserList(atacanteindex).ChatCombate = 1 Then
                    Call WriteBlockedWithShieldOther(atacanteindex)

                End If

                If UserList(victimaindex).ChatCombate = 1 Then
                    Call WriteBlockedWithShieldUser(victimaindex)

                End If

                Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 88, 0))
                
                Call SubirSkill(victimaindex, Defensa)

            End If

        End If

    End If
        
    If UsuarioImpacto Then
        If Arma > 0 Then
            If Not proyectil Then
                Call SubirSkill(atacanteindex, Armas)
            Else
                Call SubirSkill(atacanteindex, Proyectiles)

            End If

        Else
            Call SubirSkill(atacanteindex, Wrestling)

        End If

    End If

End Function

Public Sub UsuarioAtacaUsuario(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)

    Dim Probabilidad As Byte

    Dim HuboEfecto   As Boolean
    
    If Not PuedeAtacar(atacanteindex, victimaindex) Then Exit Sub
    
    If Distancia(UserList(atacanteindex).Pos, UserList(victimaindex).Pos) > MAXDISTANCIAARCO Then
        Call WriteLocaleMsg(atacanteindex, "8", FontTypeNames.FONTTYPE_INFO)
        ' Call WriteConsoleMsg(atacanteindex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub

    End If

    HuboEfecto = False
    
    Call UsuarioAtacadoPorUsuario(atacanteindex, victimaindex)
    
    If UsuarioImpacto(atacanteindex, victimaindex) Then
        Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessagePlayWave(SND_IMPACTO, UserList(atacanteindex).Pos.x, UserList(atacanteindex).Pos.Y))
        
        If UserList(victimaindex).flags.Navegando = 0 Or UserList(victimaindex).flags.Montado = 0 Then
            Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, FXSANGRE, 0))

        End If
        
        If UserList(atacanteindex).flags.incinera = 1 Then
            Probabilidad = RandomNumber(1, 6)

            If Probabilidad = 1 Then
                If UserList(victimaindex).flags.Incinerado = 0 Then
                    UserList(victimaindex).flags.Incinerado = 1

                    If UserList(victimaindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    If UserList(atacanteindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    HuboEfecto = True
                    Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Incinerar, 100, False))

                End If

            End If

        End If
    
        If UserList(atacanteindex).flags.Envenena > 0 And Not HuboEfecto Then
            Probabilidad = RandomNumber(1, 2)
    
            If Probabilidad = 1 Then
                If UserList(victimaindex).flags.Envenenado = 0 Then
                    UserList(victimaindex).flags.Envenenado = UserList(atacanteindex).flags.Envenena

                    If UserList(victimaindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If
                    
                    If UserList(atacanteindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Envenena, 100, False))

                End If

            End If

        End If
        
        If UserList(atacanteindex).flags.Paraliza = 1 And Not HuboEfecto Then
            Probabilidad = RandomNumber(1, 5)

            If Probabilidad = 1 Then
                If UserList(victimaindex).flags.Paralizado = 0 Then
                    UserList(victimaindex).flags.Paralizado = 1
                    UserList(victimaindex).Counters.Paralisis = 6
                    Call WriteParalizeOK(victimaindex)
                    Rem   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
                    If UserList(victimaindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha paralizado!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If
                    
                    If UserList(atacanteindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(atacanteindex, "Has paralizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    'Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, ParticulasIndex.Paralizar, 100, False))
                    Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 8, 0))

                End If

            End If

        End If
        
        Call UserDañoUser(atacanteindex, victimaindex)

    Else
        Call SendData(SendTarget.ToPCArea, atacanteindex, PrepareMessageCharSwing(UserList(atacanteindex).Char.CharIndex))

    End If

End Sub

Public Sub UserDañoUser(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)

    Dim daño As Long, antdaño As Integer

    Dim Lugar    As Integer, absorbido As Long

    Dim defbarco As Integer

    Dim apudaño As Integer

    Dim obj As ObjData
    
    daño = CalcularDaño(atacanteindex)
    antdaño = daño

    If PuedeApuñalar(atacanteindex) Then
        apudaño = ApuñalarFunction(atacanteindex, 0, victimaindex, daño)
        daño = daño + apudaño
        antdaño = daño

    End If

    Call UserDañoEspecial(atacanteindex, victimaindex)
    
    If UserList(atacanteindex).flags.Navegando = 1 And UserList(atacanteindex).Invent.BarcoObjIndex > 0 Then
        obj = ObjData(UserList(atacanteindex).Invent.BarcoObjIndex)
        daño = daño + RandomNumber(obj.MinHIT, obj.MaxHit)

    End If
    
    If UserList(victimaindex).flags.Navegando = 1 And UserList(victimaindex).Invent.BarcoObjIndex > 0 Then
        obj = ObjData(UserList(victimaindex).Invent.BarcoObjIndex)
        defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

    End If
    
    If UserList(atacanteindex).flags.Montado = 1 And UserList(atacanteindex).Invent.MonturaObjIndex > 0 Then
        obj = ObjData(UserList(atacanteindex).Invent.MonturaObjIndex)
        daño = daño + RandomNumber(obj.MinHIT, obj.MaxHit)

    End If
    
    If UserList(victimaindex).flags.Montado = 1 And UserList(victimaindex).Invent.MonturaObjIndex > 0 Then
        obj = ObjData(UserList(victimaindex).Invent.MonturaObjIndex)
        defbarco = RandomNumber(obj.MinDef, obj.MaxDef)

    End If
    
    Dim Resist As Byte

    If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
        Resist = ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).Refuerzo

    End If
    
    Lugar = RandomNumber(1, 6)
    
    Select Case Lugar
      
        Case PartesCuerpo.bCabeza

            'Si tiene casco absorbe el golpe
            If UserList(victimaindex).Invent.CascoEqpObjIndex > 0 Then
                obj = ObjData(UserList(victimaindex).Invent.CascoEqpObjIndex)
                absorbido = RandomNumber(obj.MinDef, obj.MaxDef)
                absorbido = absorbido + defbarco - Resist
                daño = daño - absorbido

                If daño < 0 Then daño = 1

            End If

        Case Else

            'Si tiene armadura absorbe el golpe
            If UserList(victimaindex).Invent.ArmourEqpObjIndex > 0 Then
                obj = ObjData(UserList(victimaindex).Invent.ArmourEqpObjIndex)

                Dim Obj2 As ObjData

                If UserList(victimaindex).Invent.EscudoEqpObjIndex Then
                    Obj2 = ObjData(UserList(victimaindex).Invent.EscudoEqpObjIndex)
                    absorbido = RandomNumber(obj.MinDef + Obj2.MinDef, obj.MaxDef + Obj2.MaxDef)
                Else
                    absorbido = RandomNumber(obj.MinDef, obj.MaxDef)

                End If

                absorbido = absorbido + defbarco - Resist
                daño = daño - absorbido

                If daño < 0 Then daño = 1

            End If

    End Select
    
    If apudaño > 0 Then
        Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEfectOverHead("¡" & daño & "!", UserList(victimaindex).Char.CharIndex, &HFFFF00))
    Else
        Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageEfectOverHead(daño, UserList(victimaindex).Char.CharIndex))

    End If
    
    If UserList(atacanteindex).ChatCombate = 1 Then
        Call WriteUserHittedUser(atacanteindex, Lugar, UserList(victimaindex).Char.CharIndex, daño - apudaño)

    End If
    
    If UserList(victimaindex).ChatCombate = 1 Then
        Call WriteUserHittedByUser(victimaindex, Lugar, UserList(atacanteindex).Char.CharIndex, daño - apudaño)

    End If

    UserList(victimaindex).Stats.MinHp = UserList(victimaindex).Stats.MinHp - daño
    
    If UserList(atacanteindex).flags.Hambre = 0 And UserList(atacanteindex).flags.Sed = 0 Then

        'Si usa un arma quizas suba "Combate con armas"
        If UserList(atacanteindex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(UserList(atacanteindex).Invent.WeaponEqpObjIndex).proyectil Then
                'es un Arco. Sube Armas a Distancia
                Call SubirSkill(atacanteindex, Proyectiles)
            Else
                'Sube combate con armas.
                Call SubirSkill(atacanteindex, Armas)

            End If

        Else
            'sino tal vez lucha libre
            Call SubirSkill(atacanteindex, Wrestling)

        End If
            
        Call SubirSkill(victimaindex, Tacticas)

        If PuedeApuñalar(atacanteindex) Then
            Call SubirSkill(atacanteindex, Apuñalar)

        End If
    
        'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
        ' Call DoGolpeCritico(atacanteindex, 0, victimaindex, daño)
    End If
    
    If UserList(victimaindex).Stats.MinHp <= 0 Then
    
        'Store it!
        Call Statistics.StoreFrag(atacanteindex, victimaindex)
        
        Call ContarMuerte(victimaindex, atacanteindex)
    
        Call ActStats(victimaindex, atacanteindex)
    Else
        'Está vivo - Actualizamos el HP
    
        Call WriteUpdateHP(victimaindex)

    End If
    
    'Controla el nivel del usuario
    Call CheckUserLevel(atacanteindex)
    
    Call FlushBuffer(victimaindex)

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 10/01/08
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '***************************************************

    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteLocaleMsg(VictimIndex, "123", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageParticleFX(UserList(VictimIndex).Char.CharIndex, UserList(VictimIndex).Char.ParticulaFx, 0, True))
        UserList(VictimIndex).Char.ParticulaFx = 0
        
    End If
    
    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Byte
    
    UserList(VictimIndex).Counters.TiempoDeMapeo = 3
    UserList(attackerIndex).Counters.TiempoDeMapeo = 3
    
    If Status(attackerIndex) = 1 And Status(VictimIndex) = 1 Or Status(VictimIndex) = 3 Then
        Call VolverCriminal(attackerIndex)

    End If
    
    EraCriminal = Status(attackerIndex)
    
    If EraCriminal = 2 And Status(attackerIndex) < 2 Then
        Call RefreshCharStatus(attackerIndex)
    ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
        Call RefreshCharStatus(attackerIndex)

    End If

    If Status(attackerIndex) = 2 Then If UserList(attackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(attackerIndex)
    
    'If UserList(VictimIndex).Familiar.Existe = 1 Then
    '  If UserList(VictimIndex).Familiar.Invocado = 1 Then
    '  Npclist(UserList(VictimIndex).Familiar.Id).flags.AttackedBy = UserList(attackerIndex).name
    '  Npclist(UserList(VictimIndex).Familiar.Id).Movement = TipoAI.NPCDEFENSA
    '  Npclist(UserList(VictimIndex).Familiar.Id).Hostile = 1
    ' End If
    ' End If
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)

End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

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
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(attackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function

    End If

    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(attackerIndex, "No podés atacar a un espiritu.", FontTypeNames.FONTTYPE_INFOIAO)
        PuedeAtacar = False
        Exit Function

    End If

    If UserList(attackerIndex).flags.Maldicion = 1 Then
        Call WriteConsoleMsg(attackerIndex, "¡Estas maldito! No podes atacar.", FontTypeNames.FONTTYPE_INFOIAO)
        PuedeAtacar = False
        Exit Function

    End If

    'Es miembro del grupo?
    If UserList(attackerIndex).Grupo.EnGrupo = True Then

        Dim i As Byte

        For i = 1 To UserList(UserList(attackerIndex).Grupo.Lider).Grupo.CantidadMiembros
    
            If UserList(UserList(attackerIndex).Grupo.Lider).Grupo.Miembros(i) = VictimIndex Then
                PuedeAtacar = False
                Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Function

            End If

        Next i

    End If

    'Estamos en una Arena? o un trigger zona segura?
    T = TriggerZonaPelea(attackerIndex, VictimIndex)

    If T = eTrigger6.TRIGGER6_PERMITE Then
        PuedeAtacar = True
        Exit Function
    ElseIf T = eTrigger6.TRIGGER6_PROHIBE Then
        PuedeAtacar = False
        Exit Function
    ElseIf T = eTrigger6.TRIGGER6_AUSENTE Then

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

    If UserList(attackerIndex).GuildIndex <> 0 Then
        If UserList(attackerIndex).flags.SeguroClan Then
            If UserList(attackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
                Call WriteConsoleMsg(attackerIndex, "No podes atacar a un miembro de tu clan. Para hacerlo debes desactivar el seguro de clan.", FontTypeNames.FONTTYPE_INFOIAO)
                PuedeAtacar = False
                Exit Function

            End If

        End If

    End If

    'Estas queriendo atacar a un GM?
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

    If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
        If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(attackerIndex, "El ser es demasiado poderoso", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If

    'Sos un Armada atacando un ciudadano?
    If (Status(VictimIndex) = 1) And (esArmada(attackerIndex)) Then
        Call WriteConsoleMsg(attackerIndex, "Los soldados del Ejercito Real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If

    'Tenes puesto el seguro?
    If UserList(attackerIndex).flags.Seguro Then
        If Status(VictimIndex) = 1 Then
            Call WriteConsoleMsg(attackerIndex, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function

        End If

    End If

    'Es un ciuda queriando atacar un imperial?

    If UserList(attackerIndex).flags.Seguro Then
        If (Status(attackerIndex) = 1) And (esArmada(VictimIndex)) Then
            Call WriteConsoleMsg(attackerIndex, "Los ciudadanos no pueden atacar a los soldados imperiales.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function

        End If

    End If

    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).Pos.Map).Seguro = 1 Then
        If esArmada(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasReal > 11 Then
                If UserList(VictimIndex).Pos.Map = 58 Or UserList(VictimIndex).Pos.Map = 59 Or UserList(VictimIndex).Pos.Map = 60 Then
                    Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                    Exit Function

                End If

            End If

        End If

        If esCaos(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasCaos > 11 Then
                If UserList(VictimIndex).Pos.Map = 151 Or UserList(VictimIndex).Pos.Map = 156 Then
                    Call WriteConsoleMsg(VictimIndex, "Huye de la ciudad! estas siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                    Exit Function

                End If

            End If

        End If

        Call WriteConsoleMsg(attackerIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If

    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.x, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(attackerIndex).Pos.Map, UserList(attackerIndex).Pos.x, UserList(attackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(attackerIndex, "No podes pelear aqui.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function

    End If

    PuedeAtacar = True

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

    'Estas muerto?
    If UserList(attackerIndex).flags.Muerto = 1 Then
        'Call WriteConsoleMsg(attackerIndex, "No podés atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
        Call WriteLocaleMsg(attackerIndex, "77", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacarNPC = False
        Exit Function

    End If

    'Sos consejero?
    If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
        'No pueden atacar NPC los Consejeros.
        PuedeAtacarNPC = False
        Exit Function

    End If

    'Es una criatura atacable?
    If Npclist(NpcIndex).Attackable = 0 Then
        'No es una criatura atacable
        Call WriteConsoleMsg(attackerIndex, "No podés atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacarNPC = False
        Exit Function

    End If

    'Es valida la distancia a la cual estamos atacando?
    If Distancia(UserList(attackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
        Call WriteLocaleMsg(attackerIndex, "8", FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(attackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function

    End If

    'Es una criatura No-Hostil?
    If Npclist(NpcIndex).Hostile = 0 Then
        'Es una criatura No-Hostil.
        'Es Guardia del Caos?

        If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then

            'Lo quiere atacar un caos?
            If esCaos(attackerIndex) Then
                Call WriteConsoleMsg(attackerIndex, "No podés atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function

            End If

            If Status(attackerIndex) = 1 Then
                PuedeAtacarNPC = True
                Exit Function

            End If
        
        End If

        'Es guardia Real?
        If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            'Lo quiere atacar un Armada?
        
            If esCaos(attackerIndex) Then
                PuedeAtacarNPC = True
                Exit Function

            End If
        
            If esArmada(attackerIndex) Then
                Call WriteConsoleMsg(attackerIndex, "No podés atacar Guardias Reales siendo Armada Real", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function

            End If
        
            'Tienes el seguro puesto?
            If UserList(attackerIndex).flags.Seguro And Status(attackerIndex) = 1 Then
                Call WriteConsoleMsg(attackerIndex, "Debes quitar el seguro para poder Atacar Guardias Reales utilizando /seg", FontTypeNames.FONTTYPE_INFO)
                PuedeAtacarNPC = False
                Exit Function
            Else
                Call WriteConsoleMsg(attackerIndex, "Atacaste un Guardia Real! Eres un Criminal.", FontTypeNames.FONTTYPE_INFO)
                Call VolverCriminal(attackerIndex)
                PuedeAtacarNPC = True
                Exit Function

            End If

        End If

    End If

    PuedeAtacarNPC = True

End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/09/06 Nacho
    'Reescribi gran parte del Sub
    'Ahora, da toda la experiencia del npc mientras este vivo.
    '***************************************************

    If UserList(UserIndex).Grupo.EnGrupo Then
        Call CalcularDarExpGrupal(UserIndex, NpcIndex, ElDaño)
    Else

        Dim ExpaDar As Long
    
        '[Nacho] Chekeamos que las variables sean validas para las operaciones
        If ElDaño <= 0 Then ElDaño = 0
        If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
        If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
    
        '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
        ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

        If ExpaDar <= 0 Then Exit Sub

        '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
        If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
            ExpaDar = Npclist(NpcIndex).flags.ExpCount
            Npclist(NpcIndex).flags.ExpCount = 0
        Else
            Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar

        End If
    
        If ExpMult > 0 Then
            ExpaDar = ExpaDar * ExpMult * UserList(UserIndex).flags.ScrollExp
    
        End If
    
        If UserList(UserIndex).donador.activo = 1 Then
            ExpaDar = ExpaDar * 1.1

        End If
    
        '[Nacho] Le damos la exp al user
        If ExpaDar > 0 Then
            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)

            End If
            
            Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, ExpaDar, 6)

        End If

    End If

End Sub

Sub CalcularDarExpGrupal(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)

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
    If ElDaño <= 0 Then ElDaño = 0
    If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
    If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
    
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng((ElDaño) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))

    If ExpaDar <= 0 Then Exit Sub

    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
    'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
    'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar

    End If
    
    Select Case UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
        Case 1
            BonificacionGrupo = 1

        Case 2
            BonificacionGrupo = 1.2

        Case 3
            BonificacionGrupo = 1.4

        Case 4
            BonificacionGrupo = 1.6

        Case 5
            BonificacionGrupo = 1.8

        Case 6
            BonificacionGrupo = 2
                
    End Select
 
    If ExpMult > 0 Then
        ExpaDar = ExpaDar * ExpMult
        
    End If
    
    Dim expbackup As Long

    expbackup = ExpaDar
    ExpaDar = ExpaDar * BonificacionGrupo

    Dim i     As Byte

    Dim Index As Byte

    expbackup = expbackup / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    ExpaDar = ExpaDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
    Dim ExpUser As Long
    
    For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
        Index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)

        If UserList(Index).flags.Muerto = 0 Then
            If UserList(UserIndex).Pos.Map = UserList(Index).Pos.Map Then
                If ExpaDar > 0 Then
                    ExpUser = 0

                    If UserList(Index).donador.activo = 1 Then
                        ExpUser = ExpaDar * 1.1
                    Else
                        ExpUser = ExpaDar

                    End If
                    
                    ExpUser = ExpUser * UserList(Index).flags.ScrollExp
                
                    If UserList(Index).Stats.ELV < STAT_MAXELV Then
                        UserList(Index).Stats.Exp = UserList(Index).Stats.Exp + ExpUser

                        If UserList(Index).Stats.Exp > MAXEXP Then UserList(Index).Stats.Exp = MAXEXP

                        If UserList(Index).ChatCombate = 1 Then
                            Call WriteLocaleMsg(Index, "141", FontTypeNames.FONTTYPE_EXP, ExpUser)

                        End If

                        Call WriteUpdateExp(Index)
                        Call CheckUserLevel(Index)

                    End If

                End If

            Else

                'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
                If UserList(Index).ChatCombate = 1 Then
                    Call WriteLocaleMsg(Index, "69", FontTypeNames.FONTTYPE_New_GRUPO)

                End If

                If expbackup > 0 Then
                    If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + expbackup

                        If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

                        If UserList(UserIndex).ChatCombate = 1 Then
                            Call WriteConsoleMsg(UserIndex, UserList(Index).name & " estas demasiado lejos de tu grupo, has ganado " & expbackup & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)

                        End If

                        Call CheckUserLevel(UserIndex)
                        Call WriteUpdateExp(UserIndex)

                    End If

                End If

            End If

        Else

            If UserList(Index).ChatCombate = 1 Then
                Call WriteConsoleMsg(Index, "Estas muerto, no has ganado experencia del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If

            If expbackup > 0 Then
                If UserList(Index).Stats.ELV < STAT_MAXELV Then
                    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + expbackup

                    If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP

                    If UserList(UserIndex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(UserIndex, UserList(Index).name & " estas muerto, has ganado " & expbackup & " puntos de experiencia correspondientes a el.", FontTypeNames.FONTTYPE_EXP)

                    End If

                    Call CheckUserLevel(UserIndex)
                    Call WriteUpdateExp(UserIndex)

                End If

            End If

        End If

    Next i

    'Else
    '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, experencia perdida.", FontTypeNames.FONTTYPE_New_GRUPO)
    'End If

End Sub

Sub CalcularDarOroGrupal(ByVal UserIndex As Integer, ByVal GiveGold As Long)

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/09/06 Nacho
    'Reescribi gran parte del Sub
    'Ahora, da toda la experiencia del npc mientras este vivo.
    '***************************************************
    Dim OroDar            As Long

    Dim BonificacionGrupo As Single

    'If UserList(UserIndex).Grupo.EnGrupo Then
    
    Select Case UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
    
        Case 1
            BonificacionGrupo = 1

        Case 2
            BonificacionGrupo = 1.2

        Case 3
            BonificacionGrupo = 1.4

        Case 4
            BonificacionGrupo = 1.6

        Case 5
            BonificacionGrupo = 1.8

        Case 6
            BonificacionGrupo = 2
                
    End Select
 
    OroDar = GiveGold * OroMult
    
    If OroDar > 0 Then
        OroDar = OroDar * BonificacionGrupo
        
    End If
    
    Dim orobackup As Long
    
    orobackup = OroDar

    Dim i     As Byte

    Dim Index As Byte

    OroDar = OroDar / UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

    For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
        Index = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)

        If UserList(Index).flags.Muerto = 0 Then
            If UserList(UserIndex).Pos.Map = UserList(Index).Pos.Map Then
                If OroDar > 0 Then
                    
                    OroDar = orobackup * UserList(Index).flags.ScrollOro
                
                    UserList(Index).Stats.GLD = UserList(Index).Stats.GLD + OroDar
                        
                    If UserList(Index).ChatCombate = 1 Then
                        Call WriteConsoleMsg(Index, "¡El grupo ha ganado " & OroDar & " monedas de oro!", FontTypeNames.FONTTYPE_New_GRUPO)

                    End If
                        
                    Call WriteUpdateGold(Index)
                        
                End If

            Else

                'Call WriteConsoleMsg(Index, "Estas demasiado lejos del grupo, no has ganado experiencia.", FontTypeNames.FONTTYPE_INFOIAO)
                'Call WriteLocaleMsg(Index, "69", FontTypeNames.FONTTYPE_INFOIAO)
            End If

        Else

            '  Call WriteConsoleMsg(Index, "Estas muerto, no has ganado oro del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        End If

    Next i

    'Else
    '    Call WriteConsoleMsg(UserIndex, "No te encontras en ningun grupo, oro perdido.", FontTypeNames.FONTTYPE_New_GRUPO)
    'End If

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6

    'TODO: Pero que rebuscado!!
    'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
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

    Dim ArmaObjInd As Integer, ObjInd As Integer

    Dim num        As Long
 
    ArmaObjInd = UserList(atacanteindex).Invent.WeaponEqpObjIndex
    ObjInd = 0
 
    If ArmaObjInd > 0 Then
        If ObjData(ArmaObjInd).proyectil = 0 Then
            ObjInd = ArmaObjInd
        Else
            ObjInd = UserList(atacanteindex).Invent.MunicionEqpObjIndex

        End If
   
        If ObjInd > 0 Then
            If (ObjData(ObjInd).incinera = 1) Then
                num = RandomNumber(1, 6)
           
                If num < 6 Then
                    UserList(victimaindex).flags.Incinerado = 1
                    Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                End If

            End If

        End If

    End If
 
    Call FlushBuffer(victimaindex)

End Sub

Sub UserDañoEspecial(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)

    Dim ArmaObjInd As Integer, ObjInd As Integer

    Dim HuboEfecto As Boolean

    Dim num        As Long

    HuboEfecto = False
    ArmaObjInd = UserList(atacanteindex).Invent.WeaponEqpObjIndex
    ObjInd = 0

    If ArmaObjInd = 0 Then
        ArmaObjInd = UserList(atacanteindex).Invent.NudilloObjIndex

    End If

    If ArmaObjInd > 0 Then
        If ObjData(ArmaObjInd).proyectil = 0 Then
            ObjInd = ArmaObjInd
        Else
            ObjInd = UserList(atacanteindex).Invent.MunicionEqpObjIndex

        End If
    
        If ObjInd > 0 Then
            If (ObjData(ObjInd).Envenena > 0) And Not HuboEfecto Then
                num = RandomNumber(1, 100)
            
                If num < 30 Then
                    UserList(victimaindex).flags.Envenenado = ObjData(ObjInd).Envenena
                    Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                    HuboEfecto = True

                End If

            End If
        
            If (ObjData(ObjInd).incinera > 0) And Not HuboEfecto Then
                num = RandomNumber(1, 100)
            
                If num < 10 Then
                    UserList(victimaindex).flags.Incinerado = 1
                    Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha Incinerado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(atacanteindex, "Has Incinerado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                    HuboEfecto = True

                End If

            End If
        
            If (ObjData(ObjInd).Paraliza > 0) And Not HuboEfecto Then
                num = RandomNumber(1, 100)

                If num < 10 Then
                    If UserList(victimaindex).flags.Paralizado = 0 Then
                        UserList(victimaindex).flags.Paralizado = 1
                        UserList(victimaindex).Counters.Paralisis = 6
                        Call WriteParalizeOK(victimaindex)
                        Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageCreateFX(UserList(victimaindex).Char.CharIndex, 8, 0))
                    
                        If UserList(victimaindex).ChatCombate = 1 Then
                            Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha paralizado!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If
                    
                        If UserList(atacanteindex).ChatCombate = 1 Then
                            Call WriteConsoleMsg(atacanteindex, "Has paralizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                        End If

                        HuboEfecto = True
                    
                    End If

                End If

            End If
        
            If (ObjData(ObjInd).Estupidiza > 0) And Not HuboEfecto Then
                num = RandomNumber(1, 100)

                If num < 8 Then
                    If UserList(victimaindex).flags.Estupidez = 0 Then
                        UserList(victimaindex).flags.Estupidez = 1
                        UserList(victimaindex).Counters.Estupidez = 5

                    End If
                
                    If UserList(victimaindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha estupidizado!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    Call SendData(SendTarget.ToPCArea, victimaindex, PrepareMessageParticleFX(UserList(victimaindex).Char.CharIndex, 30, 30, False))
                
                    If UserList(atacanteindex).ChatCombate = 1 Then
                        Call WriteConsoleMsg(atacanteindex, "Has estupidizado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                    End If

                    Call WriteDumb(victimaindex)

                End If

            End If

        End If

    End If

    Call FlushBuffer(victimaindex)

End Sub

