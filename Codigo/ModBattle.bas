Attribute VB_Name = "ModBattle"
Public Sub AumentarPJ(ByVal UserIndex As Integer)


Dim DistVida(1 To 5) As Integer
Dim vidaOk As Integer
        Dim manaok As Integer
        Dim staok As Integer
        Dim maxhitok As Integer
        Dim minhitok As Integer
           Dim AumentoMANA As Integer
        Dim AumentoHP As Integer
        
        Dim AumentoSTA As Integer
Dim AumentoHIT As Integer
 
        Dim i As Byte
        vidaOk = UserList(UserIndex).Stats.MaxHp
        manaok = UserList(UserIndex).Stats.MaxMAN
        staok = UserList(UserIndex).Stats.MaxSta
        maxhitok = UserList(UserIndex).Stats.MaxHit
        minhitok = UserList(UserIndex).Stats.MinHIT
        
        
        
       UserList(UserIndex).flags.LevelBackup = UserList(UserIndex).Stats.ELV
        
    Dim magia As Boolean

        
        
        Dim level As Byte
        
        For i = UserList(UserIndex).Stats.ELV + 1 To 50
        
        
        
                                'Calculo subida de vida
                        Dim RESTA As Byte
                        RESTA = 22 - UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion)
                        
                        If i < 16 Then
                            AumentoHP = ModVida(UserList(UserIndex).raza).N1TO15(UserList(UserIndex).clase) - RESTA
                        ElseIf i < 36 Then
                            AumentoHP = ModVida(UserList(UserIndex).raza).N16TO35(UserList(UserIndex).clase) - RESTA
                        ElseIf i < 46 Then
                            AumentoHP = ModVida(UserList(UserIndex).raza).N36TO45(UserList(UserIndex).clase) - RESTA
                          Else
                            AumentoHP = ModVida(UserList(UserIndex).raza).N46TO50(UserList(UserIndex).clase) - RESTA
                        End If
   
            
            

        
                                Select Case UserList(UserIndex).clase
                            
                            
                                    Case eClass.Mage '
            
                                        AumentoMANA = 3.5 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                                       ' AumentoHP = RandomNumber(MagoVidaMin, MagoVidaMax)
                                        AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
                                        AumentoSTA = AumentoSTMago
                                        magia = True
                                
                                    Case eClass.Bard 'Balanceda Mana
                                        AumentoMANA = 2.6 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                                       ' AumentoHP = RandomNumber(BardoVidaMin, BardoVidaMax)
                                        magia = True
                                        AumentoHIT = 2
                                        AumentoSTA = AumentoSTDef - 4
                                        
                                    Case eClass.Druid 'Balanceda Mana
                                        AumentoMANA = 2.9 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                                      '  AumentoHP = RandomNumber(DruidaVidaMin, DruidaVidaMax)
                                        AumentoHIT = 2
                                        AumentoSTA = AumentoSTDef - 4
                                        magia = True
                                
                                    Case eClass.Assasin
                                        AumentoMANA = 1.1 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                                       ' AumentoHP = RandomNumber(AsesinoVidaMin, AsesinoVidaMax)
                                        AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 1, 3)
                                        AumentoSTA = AumentoSTDef - 3
                                        magia = True
                                    Case eClass.Cleric 'Balanceda Mana
                                        AumentoHIT = 2
                                        AumentoMANA = 2 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                                        AumentoSTA = AumentoSTDef - 4
                                       ' AumentoHP = RandomNumber(ClerigoVidaMin, ClerigoVidaMax)
                                        magia = True
                                        
                                        
                                    Case eClass.Paladin
                                        AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 39, 1, 3)
                                        AumentoMANA = 1.1 * UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia)
                                        AumentoSTA = AumentoSTDef - 2
                                       ' AumentoHP = RandomNumber(PaladinVidaMin, PaladinVidaMax)
                                        magia = True
                                        
                                    Case eClass.Hunter
                                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
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
                                    AumentoHIT = IIf(UserList(UserIndex).Stats.ELV > 35, 2, 3)
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
        UserList(UserIndex).Stats.ELV = UserList(UserIndex).Stats.ELV + 1
    Next i
                        
                            
                            
                            
                            
                                                                    'Actualizamos HitPoints
                            UserList(UserIndex).Stats.MaxHp = vidaOk
                            UserList(UserIndex).Stats.MinHp = vidaOk
                            If UserList(UserIndex).Stats.MaxHp > STAT_MAXHP Then _
                                UserList(UserIndex).Stats.MaxHp = STAT_MAXHP

                                
                                If magia = True Then
                            'Actualizamos Mana
                                UserList(UserIndex).Stats.MaxMAN = manaok
                                UserList(UserIndex).Stats.MinMAN = manaok
                                If UserList(UserIndex).Stats.MaxMAN > 9999 Then _
                                    UserList(UserIndex).Stats.MaxMAN = 9999
                                    End If
                            

                    'Actualizamos Stamina
                        UserList(UserIndex).Stats.MaxSta = staok
                        UserList(UserIndex).Stats.MinSta = staok
                        If UserList(UserIndex).Stats.MaxSta > STAT_MAXSTA Then _
                            UserList(UserIndex).Stats.MaxSta = STAT_MAXSTA


    'Actualizamos Golpe Máximo
    UserList(UserIndex).Stats.MaxHit = maxhitok
    
    
    'Actualizamos Golpe Mínimo
    UserList(UserIndex).Stats.MinHIT = minhitok
    
    
    UserList(UserIndex).Stats.GLD = 25000
    UserList(UserIndex).Stats.ELV = 50
    
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 0
        
        
        Call RevivirUsuario(UserIndex)
        
        Call WriteUpdateUserStats(UserIndex)
        
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0 'Bug reparado 27/01/13
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        UserList(UserIndex).flags.Hambre = 0 'Bug reparado 27/01/13

        Call WriteUpdateHungerAndThirst(UserIndex)
        
        
            
        
        For i = 1 To NUMSKILLS
            UserList(UserIndex).Stats.UserSkills(i) = 100
        Next i
        
    For i = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(i) = 0

    Next i

        
        
 With UserList(UserIndex).flags
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
              UserList(UserIndex).Stats.UserAtributos(LoopX) = 35
        Next
        Call WriteFYA(UserIndex)
        
        
        

        
        If UserList(UserIndex).Char.Body_Aura <> "" Then
            UserList(UserIndex).Char.Body_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Body_Aura, True, 1))
        End If
        
        
        If UserList(UserIndex).Char.Arma_Aura <> "" Then
            UserList(UserIndex).Char.Arma_Aura = ""
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, True, 2))
        End If
        
        
        If UserList(UserIndex).Char.Escudo_Aura <> "" Then
            UserList(UserIndex).Char.Escudo_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Escudo_Aura, True, 3))
        End If
        
        If UserList(UserIndex).Char.Head_Aura <> "" Then
            UserList(UserIndex).Char.Head_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Head_Aura, True, 4))
        End If
        
        If UserList(UserIndex).Char.Otra_Aura <> "" Then
            UserList(UserIndex).Char.Otra_Aura = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Otra_Aura, True, 5))
        End If
        
        
        
    With UserList(UserIndex).Char
        .CascoAnim = 0
        .FX = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .ParticulaFx = 0
    End With
    
    
     
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco
    UserList(UserIndex).Char.CascoAnim = NingunCasco
            
           
    UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
    UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
    UserList(UserIndex).Invent.CascoEqpObjIndex = 0
    UserList(UserIndex).Invent.AnilloEqpSlot = 0
    UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
    UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
    
    
    If UserList(UserIndex).flags.Montado > 0 Then
        Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)
    End If
  Dim LoopC As Byte
For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
        'Actualiza el inventario
UserList(UserIndex).Invent.Object(LoopC).ObjIndex = 0
UserList(UserIndex).Invent.Object(LoopC).Amount = 0
UserList(UserIndex).Invent.Object(LoopC).Equipped = 0
    Next LoopC


'Vestimenta
        Select Case UserList(UserIndex).clase
            Case eClass.Mage
                UserList(UserIndex).Invent.NroItems = 10
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1964 'Tunica dorada
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1147 'DM +20
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1747 'Gorro Magico +RM 20
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1330 'Anillo penumbras
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(5).Amount = 10000
                UserList(UserIndex).Invent.Object(6).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                
                UserList(UserIndex).Stats.UserHechizos(1) = 26 'Inmovilizar
                UserList(UserIndex).Stats.UserHechizos(2) = 27 'Remover Paralisis
                UserList(UserIndex).Stats.UserHechizos(3) = 52 'Rafaga Ignea
                UserList(UserIndex).Stats.UserHechizos(4) = 51 'Descarga
                UserList(UserIndex).Stats.UserHechizos(5) = 53 'Apocalipsis
                UserList(UserIndex).Stats.UserHechizos(6) = 55 'Lamento de la banshee
                UserList(UserIndex).Stats.UserHechizos(7) = 56 'Juicio Final

                
            Case eClass.Bard
            
                UserList(UserIndex).Invent.NroItems = 11
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1962 'Tunica dorada
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1732 'Gorro Magico +RM 20
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1720 ' Escudo de tortuga +1
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1825 'Nudillo Oro
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
                UserList(UserIndex).Invent.Object(5).Amount = 1 '
                UserList(UserIndex).Invent.Object(6).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 36 'Pocion Verde
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                UserList(UserIndex).Invent.Object(11).ObjIndex = 39 'Pocion Amarilla
                UserList(UserIndex).Invent.Object(11).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                UserList(UserIndex).Stats.UserHechizos(1) = 25 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 26 'Inmovilizar
                UserList(UserIndex).Stats.UserHechizos(3) = 27 'Remover Paralisis
                UserList(UserIndex).Stats.UserHechizos(4) = 51 'Descarga
                UserList(UserIndex).Stats.UserHechizos(5) = 21 'Celeridad
                UserList(UserIndex).Stats.UserHechizos(6) = 22 'Fuerza
                UserList(UserIndex).Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                UserList(UserIndex).Stats.UserHechizos(8) = 52 'Rafaga Ignea
                UserList(UserIndex).Stats.UserHechizos(9) = 122 'Palabra Mortal

            Case eClass.Druid
                UserList(UserIndex).Invent.NroItems = 11
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1960 'Tunica dorada
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1849 'Baculo Larzul
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1759 'Gorro Magico +RM 20
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1727 'Escudo de tortuga +1
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
                UserList(UserIndex).Invent.Object(5).Amount = 1 '
                UserList(UserIndex).Invent.Object(6).ObjIndex = 37 '
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 37 '
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                UserList(UserIndex).Invent.Object(11).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(11).Amount = 10000
                
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                UserList(UserIndex).Stats.UserHechizos(1) = 25 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 26 'Inmovilizar
                UserList(UserIndex).Stats.UserHechizos(3) = 27 'Remover Paralisis
                UserList(UserIndex).Stats.UserHechizos(4) = 51 'Descarga
                UserList(UserIndex).Stats.UserHechizos(5) = 21 'Celeridad
                UserList(UserIndex).Stats.UserHechizos(6) = 22 'Fuerza
                UserList(UserIndex).Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                UserList(UserIndex).Stats.UserHechizos(8) = 52 'Rafaga Ignea
                UserList(UserIndex).Stats.UserHechizos(9) = 111 'Implosion
                UserList(UserIndex).Stats.UserHechizos(10) = 113 'Implosion
                
            Case eClass.Assasin
                UserList(UserIndex).Invent.NroItems = 10
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1903 'Armadura dragón Azul
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1789 'Daga Infernal
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1711 'Escudo Leon +1
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1763 'Casco Dorado
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 37 '
                UserList(UserIndex).Invent.Object(5).Amount = 10000
                UserList(UserIndex).Invent.Object(6).ObjIndex = 37 '
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                UserList(UserIndex).Stats.UserHechizos(1) = 25 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 26 'Inmovilizar
                UserList(UserIndex).Stats.UserHechizos(3) = 27 'Remover Paralisis
                UserList(UserIndex).Stats.UserHechizos(4) = 51 'Descarga
                UserList(UserIndex).Stats.UserHechizos(5) = 21 'Celeridad
                UserList(UserIndex).Stats.UserHechizos(6) = 22 'Fuerza
                UserList(UserIndex).Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                UserList(UserIndex).Stats.UserHechizos(8) = 141 'Ataque Sigiloso
                
            Case eClass.Cleric
                UserList(UserIndex).Invent.NroItems = 11
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1904 'Armadura dragón blanco
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1821 'Lazurt +1
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1709 'Escudo Torre +1
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1772 'Casco Dorado
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 37 '
                UserList(UserIndex).Invent.Object(5).Amount = 10000
                UserList(UserIndex).Invent.Object(6).ObjIndex = 37 '
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                UserList(UserIndex).Invent.Object(11).ObjIndex = 1330 ' Anillo
                UserList(UserIndex).Invent.Object(11).Amount = 1
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                UserList(UserIndex).Stats.UserHechizos(1) = 25 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 26 'Inmovilizar
                UserList(UserIndex).Stats.UserHechizos(3) = 27 'Remover Paralisis
                UserList(UserIndex).Stats.UserHechizos(4) = 51 'Descarga
                UserList(UserIndex).Stats.UserHechizos(5) = 21 'Celeridad
                UserList(UserIndex).Stats.UserHechizos(6) = 22 'Fuerza
                UserList(UserIndex).Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                UserList(UserIndex).Stats.UserHechizos(8) = 52 'Rafaga Ignea
                UserList(UserIndex).Stats.UserHechizos(9) = 131 'Destierro
                UserList(UserIndex).Stats.UserHechizos(10) = 132 'Oración divina
                UserList(UserIndex).Stats.UserHechizos(11) = 133 'Plegaria

                
            Case eClass.Paladin
                UserList(UserIndex).Invent.NroItems = 10
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1906 'Armadura Dragón Negra
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1790 'Espada Saramiana
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1696 'Escudo Torre +1
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1762 'Casco legendario
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(5).Amount = 10000
                UserList(UserIndex).Invent.Object(6).ObjIndex = 37 'Pocion Azul
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 36 'Pocion Verde
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 39 'Pocion Amarilla
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                
                UserList(UserIndex).Stats.UserHechizos(1) = 25 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 26 'Inmovilizar
                UserList(UserIndex).Stats.UserHechizos(3) = 27 'Remover Paralisis
                UserList(UserIndex).Stats.UserHechizos(4) = 51 'Descarga
                UserList(UserIndex).Stats.UserHechizos(5) = 21 'Celeridad
                UserList(UserIndex).Stats.UserHechizos(6) = 22 'Fuerza
                UserList(UserIndex).Stats.UserHechizos(7) = 23 'Furia de Uhkrul
                UserList(UserIndex).Stats.UserHechizos(8) = 100 'Golpe Iracundo
                UserList(UserIndex).Stats.UserHechizos(9) = 101 'Heroismo
                
                
            Case eClass.Hunter
                UserList(UserIndex).Invent.NroItems = 11
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1907 'Armadura dragón verde
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1875 'Armadura dragón verde
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1717 'Escudo Gema (Cazador)
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1767 'Casco legendario
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 1082 'Flecha Explosiva
                UserList(UserIndex).Invent.Object(5).Amount = 10000 '
                UserList(UserIndex).Invent.Object(6).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 38 '
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                UserList(UserIndex).Invent.Object(11).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(11).Amount = 10000
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                UserList(UserIndex).Stats.UserHechizos(1) = 152 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 151 'Inmovilizar
                

            Case eClass.Warrior
                UserList(UserIndex).Invent.NroItems = 11
                UserList(UserIndex).Invent.Object(1).ObjIndex = 1908 'Armadura Dragón Legendaria
                UserList(UserIndex).Invent.Object(1).Amount = 1 '
                UserList(UserIndex).Invent.Object(2).ObjIndex = 1830 'Harbinger Kin
                UserList(UserIndex).Invent.Object(2).Amount = 1 '
                UserList(UserIndex).Invent.Object(3).ObjIndex = 1695 'Escudo Torre +1
                UserList(UserIndex).Invent.Object(3).Amount = 1 '
                UserList(UserIndex).Invent.Object(4).ObjIndex = 1768 'Casco legendario
                UserList(UserIndex).Invent.Object(4).Amount = 1 '
                UserList(UserIndex).Invent.Object(5).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(5).Amount = 10000
                UserList(UserIndex).Invent.Object(6).ObjIndex = 38 'Pocion Roja
                UserList(UserIndex).Invent.Object(6).Amount = 10000
                UserList(UserIndex).Invent.Object(7).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(7).Amount = 10000
                UserList(UserIndex).Invent.Object(8).ObjIndex = 36 '
                UserList(UserIndex).Invent.Object(8).Amount = 10000
                UserList(UserIndex).Invent.Object(9).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(9).Amount = 10000
                UserList(UserIndex).Invent.Object(10).ObjIndex = 39 '
                UserList(UserIndex).Invent.Object(10).Amount = 10000
                UserList(UserIndex).Invent.Object(11).ObjIndex = 869 '
                UserList(UserIndex).Invent.Object(11).Amount = 1
                Call EquiparInvItem(UserIndex, 1)
                Call EquiparInvItem(UserIndex, 2)
                Call EquiparInvItem(UserIndex, 3)
                Call EquiparInvItem(UserIndex, 4)
                Call EquiparInvItem(UserIndex, 11)
                UserList(UserIndex).Stats.UserHechizos(1) = 152 'Paralizar
                UserList(UserIndex).Stats.UserHechizos(2) = 151 'Inmovilizar
        End Select
    
        Call UpdateUserHechizos(True, UserIndex, 0)

        
        Call UpdateUserInv(True, UserIndex, 0)
        
        
        
        
        
        
        
        
         
        
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

