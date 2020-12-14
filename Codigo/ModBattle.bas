Attribute VB_Name = "ModBattle"
Option Explicit

Public Sub AumentarPJ(ByVal Userindex As Integer)
        
        On Error GoTo AumentarPJ_Err
        

        Dim vidaOk      As Integer

        Dim manaok      As Integer

        Dim staok       As Integer

        Dim maxhitok    As Integer

        Dim minhitok    As Integer

        Dim AumentoMANA As Integer

        Dim AumentoHP   As Integer
        
        Dim AumentoSTA  As Integer

        Dim AumentoHIT  As Integer

100     With UserList(Userindex)
 
            Dim i As Byte

102         vidaOk = .Stats.MaxHp
104         manaok = .Stats.MaxMAN
106         staok = .Stats.MaxSta
108         maxhitok = .Stats.MaxHit
110         minhitok = .Stats.MinHIT
        
112         .flags.LevelBackup = .Stats.ELV
        
            Dim magia            As Boolean
        
            Dim level            As Byte

            Dim Promedio         As Double

            Dim aux              As Integer

            Dim DistVida(1 To 5) As Integer
        
114         For i = .Stats.ELV + 1 To 50
        
                'Calculo subida de vida
116             Promedio = ModVida(.clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
118             aux = RandomNumber(0, 100)
            
120             If Promedio - Int(Promedio) = 0.5 Then
                    'Es promedio semientero
122                 DistVida(1) = DistribucionSemienteraVida(1)
124                 DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
126                 DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
128                 DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
130                 If aux <= DistVida(1) Then
132                     AumentoHP = Promedio + 1.5
134                 ElseIf aux <= DistVida(2) Then
136                     AumentoHP = Promedio + 0.5
138                 ElseIf aux <= DistVida(3) Then
140                     AumentoHP = Promedio - 0.5
                    Else
142                     AumentoHP = Promedio - 1.5

                    End If

                Else
                    'Es promedio entero
144                 DistVida(1) = DistribucionEnteraVida(1)
146                 DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
148                 DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
150                 DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
152                 DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
154                 If aux <= DistVida(1) Then
156                     AumentoHP = Promedio + 2
158                 ElseIf aux <= DistVida(2) Then
160                     AumentoHP = Promedio + 1
162                 ElseIf aux <= DistVida(3) Then
164                     AumentoHP = Promedio
166                 ElseIf aux <= DistVida(4) Then
168                     AumentoHP = Promedio - 1
                    Else
170                     AumentoHP = Promedio - 2

                    End If
                
                End If
            
172             Select Case .clase

                    Case eClass.Mage '
174                     AumentoHIT = 1
176                     AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
178                     AumentoSTA = AumentoSTMago
            
180                 Case eClass.Bard 'Balanceda Mana
182                     AumentoHIT = 2
184                     AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
186                     AumentoSTA = AumentoSTDef - 4
                    
188                 Case eClass.Druid 'Balanceda Mana
190                     AumentoHIT = 2
192                     AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
194                     AumentoSTA = AumentoSTDef - 4
            
196                 Case eClass.Assasin
198                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
200                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
202                     AumentoSTA = AumentoSTDef - 3
                    
204                 Case eClass.Cleric 'Balanceda Mana
206                     AumentoHIT = 2
208                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
210                     AumentoSTA = AumentoSTDef - 4
                    
212                 Case eClass.Paladin
214                     AumentoHIT = IIf(.Stats.ELV > 39, 1, 3)
216                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
218                     AumentoSTA = AumentoSTDef - 2
                    
220                 Case eClass.Hunter
222                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
224                     AumentoSTA = AumentoSTDef - 2
                    
226                 Case eClass.Trabajador
228                     AumentoHIT = 2
230                     AumentoSTA = AumentoSTDef + 5
            
232                 Case eClass.Warrior
234                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
236                     AumentoSTA = AumentoSTDef
                        
238                 Case Else
240                     AumentoHIT = 2
242                     AumentoSTA = AumentoSTDef

                End Select
        
244             Select Case .clase
                            
                    Case eClass.Mage '
            
246                     AumentoMANA = 3.5 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        ' AumentoHP = RandomNumber(MagoVidaMin, MagoVidaMax)
248                     AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
250                     AumentoSTA = AumentoSTMago
252                     magia = True
                                
254                 Case eClass.Bard 'Balanceda Mana
256                     AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        ' AumentoHP = RandomNumber(BardoVidaMin, BardoVidaMax)
258                     magia = True
260                     AumentoHIT = 2
262                     AumentoSTA = AumentoSTDef - 4
                                        
264                 Case eClass.Druid 'Balanceda Mana
266                     AumentoMANA = 2.9 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        '  AumentoHP = RandomNumber(DruidaVidaMin, DruidaVidaMax)
268                     AumentoHIT = 2
270                     AumentoSTA = AumentoSTDef - 4
272                     magia = True
                                
274                 Case eClass.Assasin
276                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        ' AumentoHP = RandomNumber(AsesinoVidaMin, AsesinoVidaMax)
278                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
280                     AumentoSTA = AumentoSTDef - 3
282                     magia = True

284                 Case eClass.Cleric 'Balanceda Mana
286                     AumentoHIT = 2
288                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
290                     AumentoSTA = AumentoSTDef - 4
                        ' AumentoHP = RandomNumber(ClerigoVidaMin, ClerigoVidaMax)
292                     magia = True
                                        
294                 Case eClass.Paladin
296                     AumentoHIT = IIf(.Stats.ELV > 39, 1, 3)
298                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
300                     AumentoSTA = AumentoSTDef - 2
                        ' AumentoHP = RandomNumber(PaladinVidaMin, PaladinVidaMax)
302                     magia = True
                                        
304                 Case eClass.Hunter
306                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
308                     AumentoSTA = AumentoSTDef - 2
                        '   AumentoHP = RandomNumber(CazadorVidaMin, CazadorVidaMax)
310                     manaok = 0
312                     magia = False
                                        
314                 Case eClass.Trabajador
316                     AumentoHIT = 2
318                     AumentoSTA = AumentoSTDef + 8
                        '     AumentoHP = RandomNumber(TrabajadorVidaMin, TrabajadorVidaMax)
320                     manaok = 0
322                     magia = False
                                
324                 Case eClass.Warrior
326                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
328                     AumentoSTA = AumentoSTDef
                        '    AumentoHP = RandomNumber(GuerreroVidaMin, GuerreroVidaMax)
330                     manaok = 0
332                     magia = False

                End Select
                             
334             vidaOk = vidaOk + AumentoHP
                            
336             manaok = manaok + AumentoMANA
                             
338             staok = staok + AumentoSTA
340             maxhitok = maxhitok + AumentoHIT
342             minhitok = minhitok + AumentoHIT
344             .Stats.ELV = .Stats.ELV + 1
346         Next i
                            
            'Actualizamos HitPoints
348         .Stats.MaxHp = vidaOk
350         .Stats.MinHp = vidaOk

352         If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                                
354         If magia = True Then
                'Actualizamos Mana
356             .Stats.MaxMAN = manaok
358             .Stats.MinMAN = manaok

360             If .Stats.MaxMAN > 9999 Then .Stats.MaxMAN = 9999

            End If

            'Actualizamos Stamina
362         .Stats.MaxSta = staok
364         .Stats.MinSta = staok

366         If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

            'Actualizamos Golpe Máximo
368         .Stats.MaxHit = maxhitok
    
            'Actualizamos Golpe Mínimo
370         .Stats.MinHIT = minhitok
    
372         .Stats.GLD = 25000
374         .Stats.ELV = 50
    
376         .Stats.Exp = 0
378         .Stats.ELU = 0
        
380         Call RevivirUsuario(Userindex)
        
382         Call WriteUpdateUserStats(Userindex)
        
384         .Stats.MinAGU = .Stats.MaxAGU
386         .flags.Sed = 0 'Bug reparado 27/01/13
388         .Stats.MinHam = .Stats.MaxHam
390         .flags.Hambre = 0 'Bug reparado 27/01/13

392         Call WriteUpdateHungerAndThirst(Userindex)
        
394         For i = 1 To NUMSKILLS
396             .Stats.UserSkills(i) = 100
398         Next i
        
400         For i = 1 To MAXUSERHECHIZOS
402             .Stats.UserHechizos(i) = 0

404         Next i
        
406         With .flags
408             .DuracionEfecto = 0
410             .TipoPocion = 0
412             .TomoPocion = False
414             .Navegando = 0
416             .Oculto = 0
418             .Envenenado = 0
420             .invisible = 0
422             .Paralizado = 0
424             .Inmovilizado = 0
426             .CarroMineria = 0
428             .Montado = 0
430             .Incinerado = 0
432             .Paraliza = 0
434             .Envenena = 0
436             .NoPalabrasMagicas = 0
438             .NoMagiaEfeceto = 0
440             .incinera = 0
442             .Estupidiza = 0
444             .GolpeCertero = 0
446             .PendienteDelExperto = 0
448             .CarroMineria = 0
450             .PendienteDelSacrificio = 0
452             .AnilloOcultismo = 0
454             .NoDetectable = 0
456             .RegeneracionMana = 0
458             .RegeneracionHP = 0
460             .RegeneracionSta = 0
462             .Nadando = 0
464             .NecesitaOxigeno = False

            End With
    
            Dim LoopX As Integer

466         For LoopX = 1 To NUMATRIBUTOS
468             .Stats.UserAtributos(LoopX) = 35
            Next
470         Call WriteFYA(Userindex)
        
472         If .Char.Body_Aura <> "" Then
474             .Char.Body_Aura = 0
476             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, True, 1))

            End If
        
478         If .Char.Arma_Aura <> "" Then
480             .Char.Arma_Aura = ""
482             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, True, 2))

            End If
        
484         If .Char.Escudo_Aura <> "" Then
486             .Char.Escudo_Aura = 0
488             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, True, 3))

            End If
        
490         If .Char.Head_Aura <> "" Then
492             .Char.Head_Aura = 0
494             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, True, 4))

            End If
        
496         If .Char.Otra_Aura <> "" Then
498             .Char.Otra_Aura = 0
500             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, True, 5))
            End If
            
502         If .Char.Anillo_Aura <> "" Then
504             .Char.Anillo_Aura = 0
506             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, True, 6))
            End If
        
508         With .Char
510             .CascoAnim = 0
512             .FX = 0
514             .ShieldAnim = 0
516             .WeaponAnim = 0
518             .ParticulaFx = 0

            End With
     
520         .Char.WeaponAnim = NingunArma
522         .Char.ShieldAnim = NingunEscudo
524         .Char.CascoAnim = NingunCasco
526         .Char.CascoAnim = NingunCasco
           
528         .Invent.ArmourEqpObjIndex = 0
530         .Invent.WeaponEqpObjIndex = 0
532         .Invent.CascoEqpObjIndex = 0
534         .Invent.AnilloEqpSlot = 0
536         .Invent.MunicionEqpObjIndex = 0
538         .Invent.EscudoEqpObjIndex = 0
    
540         If .flags.Montado > 0 Then
542             Call DoMontar(Userindex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)

            End If

            Dim LoopC As Byte

544         For LoopC = 1 To .CurrentInventorySlots
                'Actualiza el inventario
546             .Invent.Object(LoopC).ObjIndex = 0
548             .Invent.Object(LoopC).Amount = 0
550             .Invent.Object(LoopC).Equipped = 0
552         Next LoopC

            'Vestimenta
554         Select Case .clase

                Case eClass.Mage
556                 .Invent.NroItems = 10
558                 .Invent.Object(1).ObjIndex = 1964 'Tunica dorada
560                 .Invent.Object(1).Amount = 1 '
562                 .Invent.Object(2).ObjIndex = 1147 'DM +20
564                 .Invent.Object(2).Amount = 1 '
566                 .Invent.Object(3).ObjIndex = 1747 'Gorro Magico +RM 20
568                 .Invent.Object(3).Amount = 1 '
570                 .Invent.Object(4).ObjIndex = 1330 'Anillo penumbras
572                 .Invent.Object(4).Amount = 1 '
574                 .Invent.Object(5).ObjIndex = 37 'Pocion Azul
576                 .Invent.Object(5).Amount = 10000
578                 .Invent.Object(6).ObjIndex = 37 'Pocion Azul
580                 .Invent.Object(6).Amount = 10000
582                 .Invent.Object(7).ObjIndex = 37 'Pocion Azul
584                 .Invent.Object(7).Amount = 10000
586                 .Invent.Object(8).ObjIndex = 38 'Pocion Roja
588                 .Invent.Object(8).Amount = 10000
590                 .Invent.Object(9).ObjIndex = 38 'Pocion Roja
592                 .Invent.Object(9).Amount = 10000
594                 .Invent.Object(10).ObjIndex = 38 'Pocion Roja
596                 .Invent.Object(10).Amount = 10000
598                 Call EquiparInvItem(Userindex, 1)
600                 Call EquiparInvItem(Userindex, 2)
602                 Call EquiparInvItem(Userindex, 3)
                
604                 .Stats.UserHechizos(1) = 26 'Inmovilizar
606                 .Stats.UserHechizos(2) = 27 'Remover Paralisis
608                 .Stats.UserHechizos(3) = 52 'Rafaga Ignea
610                 .Stats.UserHechizos(4) = 51 'Descarga
612                 .Stats.UserHechizos(5) = 53 'Apocalipsis
614                 .Stats.UserHechizos(6) = 55 'Lamento de la banshee
616                 .Stats.UserHechizos(7) = 56 'Juicio Final
                
618             Case eClass.Bard
            
620                 .Invent.NroItems = 11
622                 .Invent.Object(1).ObjIndex = 1962 'Tunica dorada
624                 .Invent.Object(1).Amount = 1 '
626                 .Invent.Object(2).ObjIndex = 1732 'Gorro Magico +RM 20
628                 .Invent.Object(2).Amount = 1 '
630                 .Invent.Object(3).ObjIndex = 1720 ' Escudo de tortuga +1
632                 .Invent.Object(3).Amount = 1 '
634                 .Invent.Object(4).ObjIndex = 1825 'Nudillo Oro
636                 .Invent.Object(4).Amount = 1 '
638                 .Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
640                 .Invent.Object(5).Amount = 1 '
642                 .Invent.Object(6).ObjIndex = 37 'Pocion Azul
644                 .Invent.Object(6).Amount = 10000
646                 .Invent.Object(7).ObjIndex = 37 'Pocion Azul
648                 .Invent.Object(7).Amount = 10000
650                 .Invent.Object(8).ObjIndex = 38 'Pocion Roja
652                 .Invent.Object(8).Amount = 10000
654                 .Invent.Object(9).ObjIndex = 38 'Pocion Roja
656                 .Invent.Object(9).Amount = 10000
658                 .Invent.Object(10).ObjIndex = 36 'Pocion Verde
660                 .Invent.Object(10).Amount = 10000
662                 .Invent.Object(11).ObjIndex = 39 'Pocion Amarilla
664                 .Invent.Object(11).Amount = 10000
666                 Call EquiparInvItem(Userindex, 1)
668                 Call EquiparInvItem(Userindex, 2)
670                 Call EquiparInvItem(Userindex, 3)
672                 Call EquiparInvItem(Userindex, 4)
674                 .Stats.UserHechizos(1) = 25 'Paralizar
676                 .Stats.UserHechizos(2) = 26 'Inmovilizar
678                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
680                 .Stats.UserHechizos(4) = 51 'Descarga
682                 .Stats.UserHechizos(5) = 21 'Celeridad
684                 .Stats.UserHechizos(6) = 22 'Fuerza
686                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
688                 .Stats.UserHechizos(8) = 52 'Rafaga Ignea
690                 .Stats.UserHechizos(9) = 122 'Palabra Mortal

692             Case eClass.Druid
694                 .Invent.NroItems = 11
696                 .Invent.Object(1).ObjIndex = 1960 'Tunica dorada
698                 .Invent.Object(1).Amount = 1 '
700                 .Invent.Object(2).ObjIndex = 1849 'Baculo Larzul
702                 .Invent.Object(2).Amount = 1 '
704                 .Invent.Object(3).ObjIndex = 1759 'Gorro Magico +RM 20
706                 .Invent.Object(3).Amount = 1 '
708                 .Invent.Object(4).ObjIndex = 1727 'Escudo de tortuga +1
710                 .Invent.Object(4).Amount = 1 '
712                 .Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
714                 .Invent.Object(5).Amount = 1 '
716                 .Invent.Object(6).ObjIndex = 37 '
718                 .Invent.Object(6).Amount = 10000
720                 .Invent.Object(7).ObjIndex = 37 '
722                 .Invent.Object(7).Amount = 10000
724                 .Invent.Object(8).ObjIndex = 38 '
726                 .Invent.Object(8).Amount = 10000
728                 .Invent.Object(9).ObjIndex = 38 '
730                 .Invent.Object(9).Amount = 10000
732                 .Invent.Object(10).ObjIndex = 36 '
734                 .Invent.Object(10).Amount = 10000
736                 .Invent.Object(11).ObjIndex = 39 '
738                 .Invent.Object(11).Amount = 10000
                
740                 Call EquiparInvItem(Userindex, 1)
742                 Call EquiparInvItem(Userindex, 2)
744                 Call EquiparInvItem(Userindex, 3)
746                 Call EquiparInvItem(Userindex, 4)
                
748                 .Stats.UserHechizos(1) = 25 'Paralizar
750                 .Stats.UserHechizos(2) = 26 'Inmovilizar
752                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
754                 .Stats.UserHechizos(4) = 51 'Descarga
756                 .Stats.UserHechizos(5) = 21 'Celeridad
758                 .Stats.UserHechizos(6) = 22 'Fuerza
760                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
762                 .Stats.UserHechizos(8) = 52 'Rafaga Ignea
764                 .Stats.UserHechizos(9) = 111 'Implosion
766                 .Stats.UserHechizos(10) = 113 'Implosion
                
768             Case eClass.Assasin
770                 .Invent.NroItems = 10
772                 .Invent.Object(1).ObjIndex = 1903 'Armadura dragón Azul
774                 .Invent.Object(1).Amount = 1 '
776                 .Invent.Object(2).ObjIndex = 1789 'Daga Infernal
778                 .Invent.Object(2).Amount = 1 '
780                 .Invent.Object(3).ObjIndex = 1711 'Escudo Leon +1
782                 .Invent.Object(3).Amount = 1 '
784                 .Invent.Object(4).ObjIndex = 1763 'Casco Dorado
786                 .Invent.Object(4).Amount = 1 '
788                 .Invent.Object(5).ObjIndex = 37 '
790                 .Invent.Object(5).Amount = 10000
792                 .Invent.Object(6).ObjIndex = 37 '
794                 .Invent.Object(6).Amount = 10000
796                 .Invent.Object(7).ObjIndex = 38 '
798                 .Invent.Object(7).Amount = 10000
800                 .Invent.Object(8).ObjIndex = 38 '
802                 .Invent.Object(8).Amount = 10000
804                 .Invent.Object(9).ObjIndex = 36 '
806                 .Invent.Object(9).Amount = 10000
808                 .Invent.Object(10).ObjIndex = 39 '
810                 .Invent.Object(10).Amount = 10000
812                 Call EquiparInvItem(Userindex, 1)
814                 Call EquiparInvItem(Userindex, 2)
816                 Call EquiparInvItem(Userindex, 3)
818                 Call EquiparInvItem(Userindex, 4)
                
820                 .Stats.UserHechizos(1) = 25 'Paralizar
822                 .Stats.UserHechizos(2) = 26 'Inmovilizar
824                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
826                 .Stats.UserHechizos(4) = 51 'Descarga
828                 .Stats.UserHechizos(5) = 21 'Celeridad
830                 .Stats.UserHechizos(6) = 22 'Fuerza
832                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
834                 .Stats.UserHechizos(8) = 141 'Ataque Sigiloso
                
836             Case eClass.Cleric
838                 .Invent.NroItems = 11
840                 .Invent.Object(1).ObjIndex = 1904 'Armadura dragón blanco
842                 .Invent.Object(1).Amount = 1 '
844                 .Invent.Object(2).ObjIndex = 1821 'Lazurt +1
846                 .Invent.Object(2).Amount = 1 '
848                 .Invent.Object(3).ObjIndex = 1709 'Escudo Torre +1
850                 .Invent.Object(3).Amount = 1 '
852                 .Invent.Object(4).ObjIndex = 1772 'Casco Dorado
854                 .Invent.Object(4).Amount = 1 '
856                 .Invent.Object(5).ObjIndex = 37 '
858                 .Invent.Object(5).Amount = 10000
860                 .Invent.Object(6).ObjIndex = 37 '
862                 .Invent.Object(6).Amount = 10000
864                 .Invent.Object(7).ObjIndex = 38 '
866                 .Invent.Object(7).Amount = 10000
868                 .Invent.Object(8).ObjIndex = 38 '
870                 .Invent.Object(8).Amount = 10000
872                 .Invent.Object(9).ObjIndex = 36 '
874                 .Invent.Object(9).Amount = 10000
876                 .Invent.Object(10).ObjIndex = 39 '
878                 .Invent.Object(10).Amount = 10000
880                 .Invent.Object(11).ObjIndex = 1330 ' Anillo
882                 .Invent.Object(11).Amount = 1
884                 Call EquiparInvItem(Userindex, 1)
886                 Call EquiparInvItem(Userindex, 2)
888                 Call EquiparInvItem(Userindex, 3)
890                 Call EquiparInvItem(Userindex, 4)
                
892                 .Stats.UserHechizos(1) = 25 'Paralizar
894                 .Stats.UserHechizos(2) = 26 'Inmovilizar
896                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
898                 .Stats.UserHechizos(4) = 51 'Descarga
900                 .Stats.UserHechizos(5) = 21 'Celeridad
902                 .Stats.UserHechizos(6) = 22 'Fuerza
904                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
906                 .Stats.UserHechizos(8) = 52 'Rafaga Ignea
908                 .Stats.UserHechizos(9) = 131 'Destierro
910                 .Stats.UserHechizos(10) = 132 'Oración divina
912                 .Stats.UserHechizos(11) = 133 'Plegaria
                
914             Case eClass.Paladin
916                 .Invent.NroItems = 10
918                 .Invent.Object(1).ObjIndex = 1906 'Armadura Dragón Negra
920                 .Invent.Object(1).Amount = 1 '
922                 .Invent.Object(2).ObjIndex = 1790 'Espada Saramiana
924                 .Invent.Object(2).Amount = 1 '
926                 .Invent.Object(3).ObjIndex = 1696 'Escudo Torre +1
928                 .Invent.Object(3).Amount = 1 '
930                 .Invent.Object(4).ObjIndex = 1762 'Casco legendario
932                 .Invent.Object(4).Amount = 1 '
934                 .Invent.Object(5).ObjIndex = 37 'Pocion Azul
936                 .Invent.Object(5).Amount = 10000
938                 .Invent.Object(6).ObjIndex = 37 'Pocion Azul
940                 .Invent.Object(6).Amount = 10000
942                 .Invent.Object(7).ObjIndex = 38 'Pocion Roja
944                 .Invent.Object(7).Amount = 10000
946                 .Invent.Object(8).ObjIndex = 38 'Pocion Roja
948                 .Invent.Object(8).Amount = 10000
950                 .Invent.Object(9).ObjIndex = 36 'Pocion Verde
952                 .Invent.Object(9).Amount = 10000
954                 .Invent.Object(10).ObjIndex = 39 'Pocion Amarilla
956                 .Invent.Object(10).Amount = 10000
958                 Call EquiparInvItem(Userindex, 1)
960                 Call EquiparInvItem(Userindex, 2)
962                 Call EquiparInvItem(Userindex, 3)
964                 Call EquiparInvItem(Userindex, 4)
                
966                 .Stats.UserHechizos(1) = 25 'Paralizar
968                 .Stats.UserHechizos(2) = 26 'Inmovilizar
970                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
972                 .Stats.UserHechizos(4) = 51 'Descarga
974                 .Stats.UserHechizos(5) = 21 'Celeridad
976                 .Stats.UserHechizos(6) = 22 'Fuerza
978                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
980                 .Stats.UserHechizos(8) = 100 'Golpe Iracundo
982                 .Stats.UserHechizos(9) = 101 'Heroismo
                
984             Case eClass.Hunter
986                 .Invent.NroItems = 11
988                 .Invent.Object(1).ObjIndex = 1907 'Armadura dragón verde
990                 .Invent.Object(1).Amount = 1 '
992                 .Invent.Object(2).ObjIndex = 1875 'Armadura dragón verde
994                 .Invent.Object(2).Amount = 1 '
996                 .Invent.Object(3).ObjIndex = 1717 'Escudo Gema (Cazador)
998                 .Invent.Object(3).Amount = 1 '
1000                 .Invent.Object(4).ObjIndex = 1767 'Casco legendario
1002                 .Invent.Object(4).Amount = 1 '
1004                 .Invent.Object(5).ObjIndex = 1082 'Flecha Explosiva
1006                 .Invent.Object(5).Amount = 10000 '
1008                 .Invent.Object(6).ObjIndex = 38 '
1010                 .Invent.Object(6).Amount = 10000
1012                 .Invent.Object(7).ObjIndex = 38 '
1014                 .Invent.Object(7).Amount = 10000
1016                 .Invent.Object(8).ObjIndex = 36 '
1018                 .Invent.Object(8).Amount = 10000
1020                 .Invent.Object(9).ObjIndex = 36 '
1022                 .Invent.Object(9).Amount = 10000
1024                 .Invent.Object(10).ObjIndex = 39 '
1026                 .Invent.Object(10).Amount = 10000
1028                 .Invent.Object(11).ObjIndex = 39 '
1030                 .Invent.Object(11).Amount = 10000
1032                 Call EquiparInvItem(Userindex, 1)
1034                 Call EquiparInvItem(Userindex, 2)
1036                 Call EquiparInvItem(Userindex, 3)
1038                 Call EquiparInvItem(Userindex, 4)
1040                 .Stats.UserHechizos(1) = 152 'Paralizar
1042                 .Stats.UserHechizos(2) = 151 'Inmovilizar

1044             Case eClass.Warrior
1046                 .Invent.NroItems = 11
1048                 .Invent.Object(1).ObjIndex = 1908 'Armadura Dragón Legendaria
1050                 .Invent.Object(1).Amount = 1 '
1052                 .Invent.Object(2).ObjIndex = 1830 'Harbinger Kin
1054                 .Invent.Object(2).Amount = 1 '
1056                 .Invent.Object(3).ObjIndex = 1695 'Escudo Torre +1
1058                 .Invent.Object(3).Amount = 1 '
1060                 .Invent.Object(4).ObjIndex = 1768 'Casco legendario
1062                 .Invent.Object(4).Amount = 1 '
1064                 .Invent.Object(5).ObjIndex = 38 'Pocion Roja
1066                 .Invent.Object(5).Amount = 10000
1068                 .Invent.Object(6).ObjIndex = 38 'Pocion Roja
1070                 .Invent.Object(6).Amount = 10000
1072                 .Invent.Object(7).ObjIndex = 36 '
1074                 .Invent.Object(7).Amount = 10000
1076                 .Invent.Object(8).ObjIndex = 36 '
1078                 .Invent.Object(8).Amount = 10000
1080                 .Invent.Object(9).ObjIndex = 39 '
1082                 .Invent.Object(9).Amount = 10000
1084                 .Invent.Object(10).ObjIndex = 39 '
1086                 .Invent.Object(10).Amount = 10000
1088                 .Invent.Object(11).ObjIndex = 869 '
1090                 .Invent.Object(11).Amount = 1
1092                 Call EquiparInvItem(Userindex, 1)
1094                 Call EquiparInvItem(Userindex, 2)
1096                 Call EquiparInvItem(Userindex, 3)
1098                 Call EquiparInvItem(Userindex, 4)
1100                 Call EquiparInvItem(Userindex, 11)
1102                 .Stats.UserHechizos(1) = 152 'Paralizar
1104                 .Stats.UserHechizos(2) = 151 'Inmovilizar

              End Select
    
1106         Call UpdateUserHechizos(True, Userindex, 0)
        
1108         Call UpdateUserInv(True, Userindex, 0)
        
          End With
        
         
          Exit Sub

AumentarPJ_Err:
1110      Call RegistrarError(Err.Number, Err.description, "ModBattle.AumentarPJ", Erl)
1112      Resume Next
         
End Sub

Sub RelogearUser(ByVal Userindex As Integer, ByRef name As String, ByRef UserCuenta As String)

        On Error GoTo ErrHandler

        'Reseteamos los FLAGS
100     UserList(Userindex).flags.Escondido = 0
102     UserList(Userindex).flags.TargetNPC = 0
104     UserList(Userindex).flags.TargetNpcTipo = eNPCType.Comun
106     UserList(Userindex).flags.TargetObj = 0
108     UserList(Userindex).flags.TargetUser = 0
110     UserList(Userindex).Char.FX = 0

        'Cargamos el personaje
        Dim Leer As New clsIniReader

112     Call Leer.Initialize(CharPath & UCase$(name) & ".chr")

        'Cargamos los datos del personaje
114     Call LoadUserInit(Userindex, Leer)

116     Call LoadUserStats(Userindex, Leer)

118     Set Leer = Nothing

120     If UserList(Userindex).Invent.EscudoEqpSlot = 0 Then UserList(Userindex).Char.ShieldAnim = NingunEscudo
122     If UserList(Userindex).Invent.CascoEqpSlot = 0 Then UserList(Userindex).Char.CascoAnim = NingunCasco
124     If UserList(Userindex).Invent.WeaponEqpSlot = 0 Then UserList(Userindex).Char.WeaponAnim = NingunArma

126     Call UpdateUserInv(True, Userindex, 0)
128     Call UpdateUserHechizos(True, Userindex, 0)

130     If UserList(Userindex).Correo.NoLeidos > 0 Then
132         Call WriteCorreoPicOn(Userindex)

        End If

134     If UserList(Userindex).flags.Paralizado Then
136         Call WriteParalizeOK(Userindex)

        End If

138     If UserList(Userindex).flags.Inmovilizado Then
140         Call WriteInmovilizaOK(Userindex)

        End If

        ''
        'TODO : Feo, esto tiene que ser parche cliente
142     If UserList(Userindex).flags.Estupidez = 0 Then
144         Call WriteDumbNoMore(Userindex)

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
146         UserList(Userindex).LogOnTime = Now
        #End If

148     UserList(Userindex).Char.speeding = VelocidadNormal
150     Call WriteVelocidadToggle(Userindex)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex))

        ''[/el oso]

        'LADDER NO SE SI QUEDO...
        'Call WriteErrorMsg(UserIndex, "LLegue 4")
152     Call WriteUpdateUserStats(Userindex)

154     Call WriteUpdateHungerAndThirst(Userindex)

        'Actualiza el Num de usuarios
        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!

156     Call WriteFYA(Userindex)

158     If UserList(Userindex).flags.Montado = 1 Then
160         UserList(Userindex).Char.speeding = VelocidadMontura
162         Call WriteEquiteToggle(Userindex)
    
        End If

164     If Status(Userindex) = 2 Or Status(Userindex) = 0 Then
166         Call WriteSafeModeOff(Userindex)
168         UserList(Userindex).flags.Seguro = False
        Else
170         UserList(Userindex).flags.Seguro = True
172         Call WriteSafeModeOn(Userindex)

        End If

        'Call modGuilds.SendGuildNews(UserIndex)

        'Load the user statistics
        'Call Statistics.UserConnected(UserIndex)

        'Call MostrarNumUsers

    

        'Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).name) & ".chr")

174     UserList(Userindex).flags.BattleModo = 0

        Exit Sub

ErrHandler:
176     Call WriteShowMessageBox(Userindex, "El personaje contiene un error, comuniquese con un miembro del staff.")
    

        'N = FreeFile
        'Log
        'Open App.Path & "\logs\Connect.log" For Append Shared As #N
        'Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
        'Close #N

End Sub

