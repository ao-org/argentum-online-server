Attribute VB_Name = "ModBattle"
Option Explicit

Public Sub AumentarPJ(ByVal UserIndex As Integer)

    Dim vidaOk      As Integer

    Dim manaok      As Integer

    Dim staok       As Integer

    Dim maxhitok    As Integer

    Dim minhitok    As Integer

    Dim AumentoMANA As Integer

    Dim AumentoHP   As Integer
        
    Dim AumentoSTA  As Integer

    Dim AumentoHIT  As Integer

    With UserList(UserIndex)
 
        Dim i As Byte

        vidaOk = .Stats.MaxHp
        manaok = .Stats.MaxMAN
        staok = .Stats.MaxSta
        maxhitok = .Stats.MaxHit
        minhitok = .Stats.MinHIT
        
        .flags.LevelBackup = .Stats.ELV
        
        Dim magia            As Boolean
        
        Dim level            As Byte

        Dim Promedio         As Double

        Dim aux              As Integer

        Dim DistVida(1 To 5) As Integer
        
        For i = .Stats.ELV + 1 To 50
        
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
                    AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
            
                Case eClass.Bard 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 4
                    
                Case eClass.Druid 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 4
            
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 3
                    
                Case eClass.Cleric 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 4
                    
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 39, 1, 3)
                    AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 2
                    
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef - 2
                    
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
        
            Select Case .clase
                            
                Case eClass.Mage '
            
                    AumentoMANA = 3.5 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    ' AumentoHP = RandomNumber(MagoVidaMin, MagoVidaMax)
                    AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
                    AumentoSTA = AumentoSTMago
                    magia = True
                                
                Case eClass.Bard 'Balanceda Mana
                    AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    ' AumentoHP = RandomNumber(BardoVidaMin, BardoVidaMax)
                    magia = True
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef - 4
                                        
                Case eClass.Druid 'Balanceda Mana
                    AumentoMANA = 2.9 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    '  AumentoHP = RandomNumber(DruidaVidaMin, DruidaVidaMax)
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef - 4
                    magia = True
                                
                Case eClass.Assasin
                    AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    ' AumentoHP = RandomNumber(AsesinoVidaMin, AsesinoVidaMax)
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoSTA = AumentoSTDef - 3
                    magia = True

                Case eClass.Cleric 'Balanceda Mana
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 4
                    ' AumentoHP = RandomNumber(ClerigoVidaMin, ClerigoVidaMax)
                    magia = True
                                        
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 39, 1, 3)
                    AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef - 2
                    ' AumentoHP = RandomNumber(PaladinVidaMin, PaladinVidaMax)
                    magia = True
                                        
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef - 2
                    '   AumentoHP = RandomNumber(CazadorVidaMin, CazadorVidaMax)
                    manaok = 0
                    magia = False
                                        
                Case eClass.Trabajador
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef + 8
                    '     AumentoHP = RandomNumber(TrabajadorVidaMin, TrabajadorVidaMax)
                    manaok = 0
                    magia = False
                                
                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                    '    AumentoHP = RandomNumber(GuerreroVidaMin, GuerreroVidaMax)
                    manaok = 0
                    magia = False

            End Select
                             
            vidaOk = vidaOk + AumentoHP
                            
            manaok = manaok + AumentoMANA
                             
            staok = staok + AumentoSTA
            maxhitok = maxhitok + AumentoHIT
            minhitok = minhitok + AumentoHIT
            .Stats.ELV = .Stats.ELV + 1
        Next i
                            
        'Actualizamos HitPoints
        .Stats.MaxHp = vidaOk
        .Stats.MinHp = vidaOk

        If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                                
        If magia = True Then
            'Actualizamos Mana
            .Stats.MaxMAN = manaok
            .Stats.MinMAN = manaok

            If .Stats.MaxMAN > 9999 Then .Stats.MaxMAN = 9999

        End If

        'Actualizamos Stamina
        .Stats.MaxSta = staok
        .Stats.MinSta = staok

        If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

        'Actualizamos Golpe Máximo
        .Stats.MaxHit = maxhitok
    
        'Actualizamos Golpe Mínimo
        .Stats.MinHIT = minhitok
    
        .Stats.GLD = 25000
        .Stats.ELV = 50
    
        .Stats.Exp = 0
        .Stats.ELU = 0
        
        Call RevivirUsuario(UserIndex)
        
        Call WriteUpdateUserStats(UserIndex)
        
        .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0 'Bug reparado 27/01/13
        .Stats.MinHam = .Stats.MaxHam
        .flags.Hambre = 0 'Bug reparado 27/01/13

        Call WriteUpdateHungerAndThirst(UserIndex)
        
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 100
        Next i
        
        For i = 1 To MAXUSERHECHIZOS
            .Stats.UserHechizos(i) = 0

        Next i
        
        With .flags
            .DuracionEfecto = 0
            .TipoPocion = 0
            .TomoPocion = False
            .Navegando = 0
            .Oculto = 0
            .Envenenado = 0
            .invisible = 0
            .Paralizado = 0
            .Inmovilizado = 0
            .CarroMineria = 0
            .DañoMagico = 0
            .Montado = 0
            .Incinerado = 0
            .ResistenciaMagica = 0
            .Paraliza = 0
            .Envenena = 0
            .NoPalabrasMagicas = 0
            .NoMagiaEfeceto = 0
            .incinera = 0
            .Estupidiza = 0
            .GolpeCertero = 0
            .PendienteDelExperto = 0
            .CarroMineria = 0
            .DañoMagico = 0
            .PendienteDelSacrificio = 0
            .AnilloOcultismo = 0
            .NoDetectable = 0
            .RegeneracionMana = 0
            .RegeneracionHP = 0
            .RegeneracionSta = 0
            .Nadando = 0
            .NecesitaOxigeno = False

        End With
    
        Dim LoopX As Integer

        For LoopX = 1 To NUMATRIBUTOS
            .Stats.UserAtributos(LoopX) = 35
        Next
        Call WriteFYA(UserIndex)
        
        If .Char.Body_Aura <> "" Then
            .Char.Body_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, True, 1))

        End If
        
        If .Char.Arma_Aura <> "" Then
            .Char.Arma_Aura = ""
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, True, 2))

        End If
        
        If .Char.Escudo_Aura <> "" Then
            .Char.Escudo_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, True, 3))

        End If
        
        If .Char.Head_Aura <> "" Then
            .Char.Head_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, True, 4))

        End If
        
        If .Char.Otra_Aura <> "" Then
            .Char.Otra_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, True, 5))

        End If
        
        With .Char
            .CascoAnim = 0
            .FX = 0
            .ShieldAnim = 0
            .WeaponAnim = 0
            .ParticulaFx = 0

        End With
     
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
        .Char.CascoAnim = NingunCasco
           
        .Invent.ArmourEqpObjIndex = 0
        .Invent.WeaponEqpObjIndex = 0
        .Invent.CascoEqpObjIndex = 0
        .Invent.AnilloEqpSlot = 0
        .Invent.MunicionEqpObjIndex = 0
        .Invent.EscudoEqpObjIndex = 0
    
        If .flags.Montado > 0 Then
            Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)

        End If

        Dim LoopC As Byte

        For LoopC = 1 To .CurrentInventorySlots
            'Actualiza el inventario
            .Invent.Object(LoopC).ObjIndex = 0
            .Invent.Object(LoopC).Amount = 0
            .Invent.Object(LoopC).Equipped = 0
        Next LoopC

        'Vestimenta
        Select Case .clase

            Case eClass.Mage
                .Invent.NroItems = 10
                .Invent.Object(1).ObjIndex = 1964 'Tunica dorada
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1147 'DM +20
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1747 'Gorro Magico +RM 20
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1330 'Anillo penumbras
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 37 'Pocion Azul
                .Invent.Object(5).Amount = 10000
                .Invent.Object(6).ObjIndex = 37 'Pocion Azul
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 37 'Pocion Azul
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 38 'Pocion Roja
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 38 'Pocion Roja
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 38 'Pocion Roja
                .Invent.Object(10).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                
                .Stats.UserHechizos(1) = 26 'Inmovilizar
                .Stats.UserHechizos(2) = 27 'Remover Paralisis
                .Stats.UserHechizos(3) = 52 'Rafaga Ignea
                .Stats.UserHechizos(4) = 51 'Descarga
                .Stats.UserHechizos(5) = 53 'Apocalipsis
                .Stats.UserHechizos(6) = 55 'Lamento de la banshee
                .Stats.UserHechizos(7) = 56 'Juicio Final
                
            Case eClass.Bard
            
                .Invent.NroItems = 11
                .Invent.Object(1).ObjIndex = 1962 'Tunica dorada
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1732 'Gorro Magico +RM 20
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1720 ' Escudo de tortuga +1
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1825 'Nudillo Oro
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
                .Invent.Object(5).Amount = 1 '
                .Invent.Object(6).ObjIndex = 37 'Pocion Azul
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 37 'Pocion Azul
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 38 'Pocion Roja
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 38 'Pocion Roja
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 36 'Pocion Verde
                .Invent.Object(10).Amount = 10000
                .Invent.Object(11).ObjIndex = 39 'Pocion Amarilla
                .Invent.Object(11).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                .Stats.UserHechizos(1) = 25 'Paralizar
                .Stats.UserHechizos(2) = 26 'Inmovilizar
                .Stats.UserHechizos(3) = 27 'Remover Paralisis
                .Stats.UserHechizos(4) = 51 'Descarga
                .Stats.UserHechizos(5) = 21 'Celeridad
                .Stats.UserHechizos(6) = 22 'Fuerza
                .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                .Stats.UserHechizos(8) = 52 'Rafaga Ignea
                .Stats.UserHechizos(9) = 122 'Palabra Mortal

            Case eClass.Druid
                .Invent.NroItems = 11
                .Invent.Object(1).ObjIndex = 1960 'Tunica dorada
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1849 'Baculo Larzul
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1759 'Gorro Magico +RM 20
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1727 'Escudo de tortuga +1
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
                .Invent.Object(5).Amount = 1 '
                .Invent.Object(6).ObjIndex = 37 '
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 37 '
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 38 '
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 38 '
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 36 '
                .Invent.Object(10).Amount = 10000
                .Invent.Object(11).ObjIndex = 39 '
                .Invent.Object(11).Amount = 10000
                
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                .Stats.UserHechizos(1) = 25 'Paralizar
                .Stats.UserHechizos(2) = 26 'Inmovilizar
                .Stats.UserHechizos(3) = 27 'Remover Paralisis
                .Stats.UserHechizos(4) = 51 'Descarga
                .Stats.UserHechizos(5) = 21 'Celeridad
                .Stats.UserHechizos(6) = 22 'Fuerza
                .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                .Stats.UserHechizos(8) = 52 'Rafaga Ignea
                .Stats.UserHechizos(9) = 111 'Implosion
                .Stats.UserHechizos(10) = 113 'Implosion
                
            Case eClass.Assasin
                .Invent.NroItems = 10
                .Invent.Object(1).ObjIndex = 1903 'Armadura dragón Azul
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1789 'Daga Infernal
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1711 'Escudo Leon +1
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1763 'Casco Dorado
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 37 '
                .Invent.Object(5).Amount = 10000
                .Invent.Object(6).ObjIndex = 37 '
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 38 '
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 38 '
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 36 '
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 39 '
                .Invent.Object(10).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                .Stats.UserHechizos(1) = 25 'Paralizar
                .Stats.UserHechizos(2) = 26 'Inmovilizar
                .Stats.UserHechizos(3) = 27 'Remover Paralisis
                .Stats.UserHechizos(4) = 51 'Descarga
                .Stats.UserHechizos(5) = 21 'Celeridad
                .Stats.UserHechizos(6) = 22 'Fuerza
                .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                .Stats.UserHechizos(8) = 141 'Ataque Sigiloso
                
            Case eClass.Cleric
                .Invent.NroItems = 11
                .Invent.Object(1).ObjIndex = 1904 'Armadura dragón blanco
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1821 'Lazurt +1
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1709 'Escudo Torre +1
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1772 'Casco Dorado
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 37 '
                .Invent.Object(5).Amount = 10000
                .Invent.Object(6).ObjIndex = 37 '
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 38 '
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 38 '
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 36 '
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 39 '
                .Invent.Object(10).Amount = 10000
                .Invent.Object(11).ObjIndex = 1330 ' Anillo
                .Invent.Object(11).Amount = 1
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                .Stats.UserHechizos(1) = 25 'Paralizar
                .Stats.UserHechizos(2) = 26 'Inmovilizar
                .Stats.UserHechizos(3) = 27 'Remover Paralisis
                .Stats.UserHechizos(4) = 51 'Descarga
                .Stats.UserHechizos(5) = 21 'Celeridad
                .Stats.UserHechizos(6) = 22 'Fuerza
                .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                .Stats.UserHechizos(8) = 52 'Rafaga Ignea
                .Stats.UserHechizos(9) = 131 'Destierro
                .Stats.UserHechizos(10) = 132 'Oración divina
                .Stats.UserHechizos(11) = 133 'Plegaria
                
            Case eClass.Paladin
                .Invent.NroItems = 10
                .Invent.Object(1).ObjIndex = 1906 'Armadura Dragón Negra
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1790 'Espada Saramiana
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1696 'Escudo Torre +1
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1762 'Casco legendario
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 37 'Pocion Azul
                .Invent.Object(5).Amount = 10000
                .Invent.Object(6).ObjIndex = 37 'Pocion Azul
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 38 'Pocion Roja
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 38 'Pocion Roja
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 36 'Pocion Verde
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 39 'Pocion Amarilla
                .Invent.Object(10).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                .Stats.UserHechizos(1) = 25 'Paralizar
                .Stats.UserHechizos(2) = 26 'Inmovilizar
                .Stats.UserHechizos(3) = 27 'Remover Paralisis
                .Stats.UserHechizos(4) = 51 'Descarga
                .Stats.UserHechizos(5) = 21 'Celeridad
                .Stats.UserHechizos(6) = 22 'Fuerza
                .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                .Stats.UserHechizos(8) = 100 'Golpe Iracundo
                .Stats.UserHechizos(9) = 101 'Heroismo
                
            Case eClass.Hunter
                .Invent.NroItems = 11
                .Invent.Object(1).ObjIndex = 1907 'Armadura dragón verde
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1875 'Armadura dragón verde
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1717 'Escudo Gema (Cazador)
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1767 'Casco legendario
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 1082 'Flecha Explosiva
                .Invent.Object(5).Amount = 10000 '
                .Invent.Object(6).ObjIndex = 38 '
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 38 '
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 36 '
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 36 '
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 39 '
                .Invent.Object(10).Amount = 10000
                .Invent.Object(11).ObjIndex = 39 '
                .Invent.Object(11).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                .Stats.UserHechizos(1) = 152 'Paralizar
                .Stats.UserHechizos(2) = 151 'Inmovilizar

            Case eClass.Warrior
                .Invent.NroItems = 11
                .Invent.Object(1).ObjIndex = 1908 'Armadura Dragón Legendaria
                .Invent.Object(1).Amount = 1 '
                .Invent.Object(2).ObjIndex = 1830 'Harbinger Kin
                .Invent.Object(2).Amount = 1 '
                .Invent.Object(3).ObjIndex = 1695 'Escudo Torre +1
                .Invent.Object(3).Amount = 1 '
                .Invent.Object(4).ObjIndex = 1768 'Casco legendario
                .Invent.Object(4).Amount = 1 '
                .Invent.Object(5).ObjIndex = 38 'Pocion Roja
                .Invent.Object(5).Amount = 10000
                .Invent.Object(6).ObjIndex = 38 'Pocion Roja
                .Invent.Object(6).Amount = 10000
                .Invent.Object(7).ObjIndex = 36 '
                .Invent.Object(7).Amount = 10000
                .Invent.Object(8).ObjIndex = 36 '
                .Invent.Object(8).Amount = 10000
                .Invent.Object(9).ObjIndex = 39 '
                .Invent.Object(9).Amount = 10000
                .Invent.Object(10).ObjIndex = 39 '
                .Invent.Object(10).Amount = 10000
                .Invent.Object(11).ObjIndex = 869 '
                .Invent.Object(11).Amount = 1
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                Call EquiparInvItem(UserIndex, 11)
                .Stats.UserHechizos(1) = 152 'Paralizar
                .Stats.UserHechizos(2) = 151 'Inmovilizar

        End Select
    
        Call UpdateUserHechizos(True, UserIndex, 0)
        
        Call UpdateUserInv(True, UserIndex, 0)
        
    End With
        
End Sub

Sub RelogearUser(ByVal UserIndex As Integer, ByRef name As String, ByRef UserCuenta As String)

    On Error GoTo Errhandler

    'Reseteamos los FLAGS
    UserList(UserIndex).flags.Escondido = 0
    UserList(UserIndex).flags.TargetNPC = 0
    UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
    UserList(UserIndex).flags.TargetObj = 0
    UserList(UserIndex).flags.TargetUser = 0
    UserList(UserIndex).Char.FX = 0

    'Cargamos el personaje
    Dim Leer As New clsIniReader

    Call Leer.Initialize(CharPath & UCase$(name) & ".chr")

    'Cargamos los datos del personaje
    Call LoadUserInit(UserIndex, Leer)

    Call LoadUserStats(UserIndex, Leer)

    Set Leer = Nothing

    If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
    If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma

    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)

    If UserList(UserIndex).Correo.NoLeidos > 0 Then
        Call WriteCorreoPicOn(UserIndex)

    End If

    If UserList(UserIndex).flags.Paralizado Then
        Call WriteParalizeOK(UserIndex)

    End If

    If UserList(UserIndex).flags.Inmovilizado Then
        Call WriteInmovilizaOK(UserIndex)

    End If

    ''
    'TODO : Feo, esto tiene que ser parche cliente
    If UserList(UserIndex).flags.Estupidez = 0 Then
        Call WriteDumbNoMore(UserIndex)

    End If

    'Posicion de comienzo

    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)

    Rem If UserList(UserIndex).Invent.MonturaObjIndex > 0 Then
    '    Debug.Print "tiene monutra"
    '    Dim Montura As ObjData
    '   Montura = ObjData(UserList(UserIndex).Invent.MonturaObjIndex)

    '    UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.MonturaObjIndex).Ropaje
    ' UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    ' UserList(UserIndex).Char.WeaponAnim = NingunArma
    '  UserList(UserIndex).Char.CascoAnim = NingunCasco
    '   UserList(UserIndex).flags.Montado = 1
    '   UserList(UserIndex).Char.Speeding = 1.3
    'End If

    'Call WriteErrorMsg(UserIndex, "LLegue 1")

    'Info

    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        UserList(UserIndex).LogOnTime = Now
    #End If

    UserList(UserIndex).Char.speeding = VelocidadNormal
    Call WriteVelocidadToggle(UserIndex)
    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex))

    ''[/el oso]

    'LADDER NO SE SI QUEDO...
    'Call WriteErrorMsg(UserIndex, "LLegue 4")
    Call WriteUpdateUserStats(UserIndex)

    Call WriteUpdateHungerAndThirst(UserIndex)

    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!

    Call WriteFYA(UserIndex)

    If UserList(UserIndex).flags.Montado = 1 Then
        UserList(UserIndex).Char.speeding = VelocidadMontura
        Call WriteEquiteToggle(UserIndex)
    
    End If

    If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
        Call WriteSafeModeOff(UserIndex)
        UserList(UserIndex).flags.Seguro = False
    Else
        UserList(UserIndex).flags.Seguro = True
        Call WriteSafeModeOn(UserIndex)

    End If

    'Call modGuilds.SendGuildNews(UserIndex)

    'Load the user statistics
    'Call Statistics.UserConnected(UserIndex)

    'Call MostrarNumUsers

    Call FlushBuffer(UserIndex)

    'Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).name) & ".chr")

    UserList(UserIndex).flags.BattleModo = 0

    Exit Sub

Errhandler:
    Call WriteShowMessageBox(UserIndex, "El personaje contiene un error, comuniquese con un miembro del staff.")
    Call FlushBuffer(UserIndex)

    'N = FreeFile
    'Log
    'Open App.Path & "\logs\Connect.log" For Append Shared As #N
    'Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
    'Close #N

End Sub

