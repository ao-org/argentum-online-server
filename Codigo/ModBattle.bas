Attribute VB_Name = "ModBattle"
Option Explicit

Public Sub AumentarPJ(ByVal UserIndex As Integer)
        
        On Error GoTo AumentarPJ_Err
        

        Dim vidaOk      As Integer

        Dim manaok      As Integer

        Dim staok       As Integer

        Dim maxhitok    As Integer

        Dim minhitok    As Integer

        Dim AumentoMANA As Integer

        Dim AumentoHP   As Integer
        
        Dim AumentoSta  As Integer

        Dim AumentoHIT  As Integer
        
        Dim PromedioObjetivo As Double
        
        Dim PromedioUser As Double
        
        Dim Promedio As Double
        
        ' Randomizo las vidas
100     Randomize Time

102     With UserList(UserIndex)
 
            Dim i As Byte

104         vidaOk = .Stats.MaxHp
106         manaok = .Stats.MaxMAN
108         staok = .Stats.MaxSta
110         maxhitok = .Stats.MaxHit
112         minhitok = .Stats.MinHIT
        
114         .flags.LevelBackup = .Stats.ELV
        
            Dim magia            As Boolean
        
            Dim level            As Byte

            Dim aux              As Integer

            Dim DistVida(1 To 5) As Integer
        
116         For i = .Stats.ELV + 1 To 50
        
                ' Calculo subida de vida by WyroX
                ' Obtengo el promedio según clase y constitución
118             PromedioObjetivo = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
                ' Obtengo el promedio actual del user
120             PromedioUser = CalcularPromedioVida(UserIndex)
                ' Lo modifico para compensar si está muy bajo o muy alto
122             Promedio = PromedioObjetivo + (PromedioObjetivo - PromedioUser) * DesbalancePromedioVidas
                ' Obtengo un entero al azar con más tendencia al promedio
124             AumentoHP = RandomIntBiased(PromedioObjetivo - RangoVidas, PromedioObjetivo + RangoVidas, Promedio, InfluenciaPromedioVidas)

126             Select Case .clase

                    Case eClass.Mage '
128                     AumentoHIT = 1
130                     AumentoMANA = 3 * .Stats.UserAtributos(eAtributos.Inteligencia)
132                     AumentoSta = AumentoSTMago
            
134                 Case eClass.Bard 'Balanceda Mana
136                     AumentoHIT = 2
138                     AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
140                     AumentoSta = AumentoSTDef - 4
                    
142                 Case eClass.Druid 'Balanceda Mana
144                     AumentoHIT = 2
146                     AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
148                     AumentoSta = AumentoSTDef - 4
            
150                 Case eClass.Assasin
152                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
154                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
156                     AumentoSta = AumentoSTDef - 3
                    
158                 Case eClass.Cleric 'Balanceda Mana
160                     AumentoHIT = 2
162                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
164                     AumentoSta = AumentoSTDef - 4
                    
166                 Case eClass.Paladin
168                     AumentoHIT = IIf(.Stats.ELV > 39, 1, 3)
170                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
172                     AumentoSta = AumentoSTDef - 2
                    
174                 Case eClass.Hunter
176                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
178                     AumentoSta = AumentoSTDef - 2
                    
180                 Case eClass.Trabajador
182                     AumentoHIT = 2
184                     AumentoSta = AumentoSTDef + 5
            
186                 Case eClass.Warrior
188                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
190                     AumentoSta = AumentoSTDef
                        
192                 Case Else
194                     AumentoHIT = 2
196                     AumentoSta = AumentoSTDef

                End Select
        
198             Select Case .clase
                            
                    Case eClass.Mage '
            
200                     AumentoMANA = 3.5 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        ' AumentoHP = RandomNumber(MagoVidaMin, MagoVidaMax)
202                     AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
204                     AumentoSta = AumentoSTMago
206                     magia = True
                                
208                 Case eClass.Bard 'Balanceda Mana
210                     AumentoMANA = 2.6 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        ' AumentoHP = RandomNumber(BardoVidaMin, BardoVidaMax)
212                     magia = True
214                     AumentoHIT = 2
216                     AumentoSta = AumentoSTDef - 4
                                        
218                 Case eClass.Druid 'Balanceda Mana
220                     AumentoMANA = 2.9 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        '  AumentoHP = RandomNumber(DruidaVidaMin, DruidaVidaMax)
222                     AumentoHIT = 2
224                     AumentoSta = AumentoSTDef - 4
226                     magia = True
                                
228                 Case eClass.Assasin
230                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
                        ' AumentoHP = RandomNumber(AsesinoVidaMin, AsesinoVidaMax)
232                     AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
234                     AumentoSta = AumentoSTDef - 3
236                     magia = True

238                 Case eClass.Cleric 'Balanceda Mana
240                     AumentoHIT = 2
242                     AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
244                     AumentoSta = AumentoSTDef - 4
                        ' AumentoHP = RandomNumber(ClerigoVidaMin, ClerigoVidaMax)
246                     magia = True
                                        
248                 Case eClass.Paladin
250                     AumentoHIT = IIf(.Stats.ELV > 39, 1, 3)
252                     AumentoMANA = 1.1 * .Stats.UserAtributos(eAtributos.Inteligencia)
254                     AumentoSta = AumentoSTDef - 2
                        ' AumentoHP = RandomNumber(PaladinVidaMin, PaladinVidaMax)
256                     magia = True
                                        
258                 Case eClass.Hunter
260                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
262                     AumentoSta = AumentoSTDef - 2
                        '   AumentoHP = RandomNumber(CazadorVidaMin, CazadorVidaMax)
264                     manaok = 0
266                     magia = False
                                        
268                 Case eClass.Trabajador
270                     AumentoHIT = 2
272                     AumentoSta = AumentoSTDef + 8
                        '     AumentoHP = RandomNumber(TrabajadorVidaMin, TrabajadorVidaMax)
274                     manaok = 0
276                     magia = False
                                
278                 Case eClass.Warrior
280                     AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
282                     AumentoSta = AumentoSTDef
                        '    AumentoHP = RandomNumber(GuerreroVidaMin, GuerreroVidaMax)
284                     manaok = 0
286                     magia = False

                End Select
                             
288             vidaOk = vidaOk + AumentoHP
                            
290             manaok = manaok + AumentoMANA
                             
292             staok = staok + AumentoSta
294             maxhitok = maxhitok + AumentoHIT
296             minhitok = minhitok + AumentoHIT
298             .Stats.ELV = .Stats.ELV + 1
300         Next i
                            
            'Actualizamos HitPoints
302         .Stats.MaxHp = vidaOk
304         .Stats.MinHp = vidaOk

306         If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                                
308         If magia = True Then
                'Actualizamos Mana
310             .Stats.MaxMAN = manaok
312             .Stats.MinMAN = manaok

314             If .Stats.MaxMAN > STAT_MAXMP Then .Stats.MaxMAN = STAT_MAXMP

            End If

            'Actualizamos Stamina
316         .Stats.MaxSta = staok
318         .Stats.MinSta = staok

320         If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA

            'Actualizamos Golpe Máximo
322         .Stats.MaxHit = maxhitok
    
            'Actualizamos Golpe Mínimo
324         .Stats.MinHIT = minhitok
    
326         .Stats.GLD = 25000
328         .Stats.ELV = 50
    
330         .Stats.Exp = 0
332         .Stats.ELU = 0
        
334         Call RevivirUsuario(UserIndex)
        
336         Call WriteUpdateUserStats(UserIndex)
        
338         .Stats.MinAGU = .Stats.MaxAGU
340         .flags.Sed = 0 'Bug reparado 27/01/13
342         .Stats.MinHam = .Stats.MaxHam
344         .flags.Hambre = 0 'Bug reparado 27/01/13

346         Call WriteUpdateHungerAndThirst(UserIndex)
        
348         For i = 1 To NUMSKILLS
350             .Stats.UserSkills(i) = 100
352         Next i
        
354         For i = 1 To MAXUSERHECHIZOS
356             .Stats.UserHechizos(i) = 0

358         Next i
        
360         With .flags
362             .DuracionEfecto = 0
364             .TipoPocion = 0
366             .TomoPocion = False
368             .Navegando = 0
370             .Oculto = 0
372             .Envenenado = 0
374             .invisible = 0
376             .Paralizado = 0
378             .Inmovilizado = 0
380             .CarroMineria = 0
382             .Montado = 0
384             .Incinerado = 0
386             .Paraliza = 0
388             .Envenena = 0
390             .NoPalabrasMagicas = 0
392             .NoMagiaEfeceto = 0
394             .incinera = 0
396             .Estupidiza = 0
398             .GolpeCertero = 0
400             .PendienteDelExperto = 0
402             .CarroMineria = 0
404             .PendienteDelSacrificio = 0
406             .AnilloOcultismo = 0
408             .NoDetectable = 0
410             .RegeneracionMana = 0
412             .RegeneracionHP = 0
414             .RegeneracionSta = 0
416             .Nadando = 0
418             .NecesitaOxigeno = False

            End With
    
            Dim LoopX As Integer

420         For LoopX = 1 To NUMATRIBUTOS
422             .Stats.UserAtributos(LoopX) = 35
            Next
424         Call WriteFYA(UserIndex)
        
426         If .Char.Body_Aura <> "" Then
428             .Char.Body_Aura = 0
430             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, True, 1))

            End If
        
432         If .Char.Arma_Aura <> "" Then
434             .Char.Arma_Aura = ""
436             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, True, 2))

            End If
        
438         If .Char.Escudo_Aura <> "" Then
440             .Char.Escudo_Aura = 0
442             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, True, 3))

            End If
        
444         If .Char.Head_Aura <> "" Then
446             .Char.Head_Aura = 0
448             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, True, 4))

            End If
        
450         If .Char.Otra_Aura <> "" Then
452             .Char.Otra_Aura = 0
454             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, True, 5))
            End If
            
456         If .Char.DM_Aura <> "" Then
458             .Char.DM_Aura = 0
460             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.DM_Aura, True, 6))
            End If
            
462         If .Char.RM_Aura <> "" Then
464             .Char.RM_Aura = 0
466             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.RM_Aura, True, 7))
            End If
        
468         With .Char
470             .CascoAnim = 0
472             .FX = 0
474             .ShieldAnim = 0
476             .WeaponAnim = 0
478             .ParticulaFx = 0

            End With
     
480         .Char.WeaponAnim = NingunArma
482         .Char.ShieldAnim = NingunEscudo
484         .Char.CascoAnim = NingunCasco
486         .Char.CascoAnim = NingunCasco
           
488         .Invent.ArmourEqpObjIndex = 0
490         .Invent.WeaponEqpObjIndex = 0
492         .Invent.CascoEqpObjIndex = 0
494         .Invent.DañoMagicoEqpObjIndex = 0
496         .Invent.ResistenciaEqpObjIndex = 0
498         .Invent.MunicionEqpObjIndex = 0
500         .Invent.EscudoEqpObjIndex = 0
    
502         If .flags.Montado > 0 Then
504             Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)

            End If

            Dim LoopC As Byte

506         For LoopC = 1 To .CurrentInventorySlots
                'Actualiza el inventario
508             .Invent.Object(LoopC).ObjIndex = 0
510             .Invent.Object(LoopC).Amount = 0
512             .Invent.Object(LoopC).Equipped = 0
514         Next LoopC

            'Vestimenta
516         Select Case .clase

                Case eClass.Mage
518                 .Invent.NroItems = 10
520                 .Invent.Object(1).ObjIndex = 1964 'Tunica dorada
522                 .Invent.Object(1).Amount = 1 '
524                 .Invent.Object(2).ObjIndex = 1147 'DM +20
526                 .Invent.Object(2).Amount = 1 '
528                 .Invent.Object(3).ObjIndex = 1747 'Gorro Magico +RM 20
530                 .Invent.Object(3).Amount = 1 '
532                 .Invent.Object(4).ObjIndex = 1330 'Anillo penumbras
534                 .Invent.Object(4).Amount = 1 '
536                 .Invent.Object(5).ObjIndex = 37 'Pocion Azul
538                 .Invent.Object(5).Amount = 10000
540                 .Invent.Object(6).ObjIndex = 37 'Pocion Azul
542                 .Invent.Object(6).Amount = 10000
544                 .Invent.Object(7).ObjIndex = 37 'Pocion Azul
546                 .Invent.Object(7).Amount = 10000
548                 .Invent.Object(8).ObjIndex = 38 'Pocion Roja
550                 .Invent.Object(8).Amount = 10000
552                 .Invent.Object(9).ObjIndex = 38 'Pocion Roja
554                 .Invent.Object(9).Amount = 10000
556                 .Invent.Object(10).ObjIndex = 38 'Pocion Roja
558                 .Invent.Object(10).Amount = 10000
560                 Call EquiparInvItem(UserIndex, 1)
562                 Call EquiparInvItem(UserIndex, 2)
564                 Call EquiparInvItem(UserIndex, 3)
                
566                 .Stats.UserHechizos(1) = 26 'Inmovilizar
568                 .Stats.UserHechizos(2) = 27 'Remover Paralisis
570                 .Stats.UserHechizos(3) = 52 'Rafaga Ignea
572                 .Stats.UserHechizos(4) = 51 'Descarga
574                 .Stats.UserHechizos(5) = 53 'Apocalipsis
576                 .Stats.UserHechizos(6) = 55 'Lamento de la banshee
578                 .Stats.UserHechizos(7) = 56 'Juicio Final
                
580             Case eClass.Bard
            
582                 .Invent.NroItems = 11
584                 .Invent.Object(1).ObjIndex = 1962 'Tunica dorada
586                 .Invent.Object(1).Amount = 1 '
588                 .Invent.Object(2).ObjIndex = 1732 'Gorro Magico +RM 20
590                 .Invent.Object(2).Amount = 1 '
592                 .Invent.Object(3).ObjIndex = 1720 ' Escudo de tortuga +1
594                 .Invent.Object(3).Amount = 1 '
596                 .Invent.Object(4).ObjIndex = 1825 'Nudillo Oro
598                 .Invent.Object(4).Amount = 1 '
600                 .Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
602                 .Invent.Object(5).Amount = 1 '
604                 .Invent.Object(6).ObjIndex = 37 'Pocion Azul
606                 .Invent.Object(6).Amount = 10000
608                 .Invent.Object(7).ObjIndex = 37 'Pocion Azul
610                 .Invent.Object(7).Amount = 10000
612                 .Invent.Object(8).ObjIndex = 38 'Pocion Roja
614                 .Invent.Object(8).Amount = 10000
616                 .Invent.Object(9).ObjIndex = 38 'Pocion Roja
618                 .Invent.Object(9).Amount = 10000
620                 .Invent.Object(10).ObjIndex = 36 'Pocion Verde
622                 .Invent.Object(10).Amount = 10000
624                 .Invent.Object(11).ObjIndex = 39 'Pocion Amarilla
626                 .Invent.Object(11).Amount = 10000
628                 Call EquiparInvItem(UserIndex, 1)
630                 Call EquiparInvItem(UserIndex, 2)
632                 Call EquiparInvItem(UserIndex, 3)
634                 Call EquiparInvItem(UserIndex, 4)
636                 .Stats.UserHechizos(1) = 25 'Paralizar
638                 .Stats.UserHechizos(2) = 26 'Inmovilizar
640                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
642                 .Stats.UserHechizos(4) = 51 'Descarga
644                 .Stats.UserHechizos(5) = 21 'Celeridad
646                 .Stats.UserHechizos(6) = 22 'Fuerza
648                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
650                 .Stats.UserHechizos(8) = 52 'Rafaga Ignea
652                 .Stats.UserHechizos(9) = 122 'Palabra Mortal

654             Case eClass.Druid
656                 .Invent.NroItems = 11
658                 .Invent.Object(1).ObjIndex = 1960 'Tunica dorada
660                 .Invent.Object(1).Amount = 1 '
662                 .Invent.Object(2).ObjIndex = 1849 'Baculo Larzul
664                 .Invent.Object(2).Amount = 1 '
666                 .Invent.Object(3).ObjIndex = 1759 'Gorro Magico +RM 20
668                 .Invent.Object(3).Amount = 1 '
670                 .Invent.Object(4).ObjIndex = 1727 'Escudo de tortuga +1
672                 .Invent.Object(4).Amount = 1 '
674                 .Invent.Object(5).ObjIndex = 1330 'Anillo penumbras
676                 .Invent.Object(5).Amount = 1 '
678                 .Invent.Object(6).ObjIndex = 37 '
680                 .Invent.Object(6).Amount = 10000
682                 .Invent.Object(7).ObjIndex = 37 '
684                 .Invent.Object(7).Amount = 10000
686                 .Invent.Object(8).ObjIndex = 38 '
688                 .Invent.Object(8).Amount = 10000
690                 .Invent.Object(9).ObjIndex = 38 '
692                 .Invent.Object(9).Amount = 10000
694                 .Invent.Object(10).ObjIndex = 36 '
696                 .Invent.Object(10).Amount = 10000
698                 .Invent.Object(11).ObjIndex = 39 '
700                 .Invent.Object(11).Amount = 10000
                
702                 Call EquiparInvItem(UserIndex, 1)
704                 Call EquiparInvItem(UserIndex, 2)
706                 Call EquiparInvItem(UserIndex, 3)
708                 Call EquiparInvItem(UserIndex, 4)
                
710                 .Stats.UserHechizos(1) = 25 'Paralizar
712                 .Stats.UserHechizos(2) = 26 'Inmovilizar
714                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
716                 .Stats.UserHechizos(4) = 51 'Descarga
718                 .Stats.UserHechizos(5) = 21 'Celeridad
720                 .Stats.UserHechizos(6) = 22 'Fuerza
722                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
724                 .Stats.UserHechizos(8) = 52 'Rafaga Ignea
726                 .Stats.UserHechizos(9) = 111 'Implosion
728                 .Stats.UserHechizos(10) = 113 'Implosion
                
730             Case eClass.Assasin
732                 .Invent.NroItems = 10
734                 .Invent.Object(1).ObjIndex = 1903 'Armadura dragón Azul
736                 .Invent.Object(1).Amount = 1 '
738                 .Invent.Object(2).ObjIndex = 1789 'Daga Infernal
740                 .Invent.Object(2).Amount = 1 '
742                 .Invent.Object(3).ObjIndex = 1711 'Escudo Leon +1
744                 .Invent.Object(3).Amount = 1 '
746                 .Invent.Object(4).ObjIndex = 1763 'Casco Dorado
748                 .Invent.Object(4).Amount = 1 '
750                 .Invent.Object(5).ObjIndex = 37 '
752                 .Invent.Object(5).Amount = 10000
754                 .Invent.Object(6).ObjIndex = 37 '
756                 .Invent.Object(6).Amount = 10000
758                 .Invent.Object(7).ObjIndex = 38 '
760                 .Invent.Object(7).Amount = 10000
762                 .Invent.Object(8).ObjIndex = 38 '
764                 .Invent.Object(8).Amount = 10000
766                 .Invent.Object(9).ObjIndex = 36 '
768                 .Invent.Object(9).Amount = 10000
770                 .Invent.Object(10).ObjIndex = 39 '
772                 .Invent.Object(10).Amount = 10000
774                 Call EquiparInvItem(UserIndex, 1)
776                 Call EquiparInvItem(UserIndex, 2)
778                 Call EquiparInvItem(UserIndex, 3)
780                 Call EquiparInvItem(UserIndex, 4)
                
782                 .Stats.UserHechizos(1) = 25 'Paralizar
784                 .Stats.UserHechizos(2) = 26 'Inmovilizar
786                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
788                 .Stats.UserHechizos(4) = 51 'Descarga
790                 .Stats.UserHechizos(5) = 21 'Celeridad
792                 .Stats.UserHechizos(6) = 22 'Fuerza
794                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
796                 .Stats.UserHechizos(8) = 141 'Ataque Sigiloso
                
798             Case eClass.Cleric
800                 .Invent.NroItems = 11
802                 .Invent.Object(1).ObjIndex = 1904 'Armadura dragón blanco
804                 .Invent.Object(1).Amount = 1 '
806                 .Invent.Object(2).ObjIndex = 1821 'Lazurt +1
808                 .Invent.Object(2).Amount = 1 '
810                 .Invent.Object(3).ObjIndex = 1709 'Escudo Torre +1
812                 .Invent.Object(3).Amount = 1 '
814                 .Invent.Object(4).ObjIndex = 1772 'Casco Dorado
816                 .Invent.Object(4).Amount = 1 '
818                 .Invent.Object(5).ObjIndex = 37 '
820                 .Invent.Object(5).Amount = 10000
822                 .Invent.Object(6).ObjIndex = 37 '
824                 .Invent.Object(6).Amount = 10000
826                 .Invent.Object(7).ObjIndex = 38 '
828                 .Invent.Object(7).Amount = 10000
830                 .Invent.Object(8).ObjIndex = 38 '
832                 .Invent.Object(8).Amount = 10000
834                 .Invent.Object(9).ObjIndex = 36 '
836                 .Invent.Object(9).Amount = 10000
838                 .Invent.Object(10).ObjIndex = 39 '
840                 .Invent.Object(10).Amount = 10000
842                 .Invent.Object(11).ObjIndex = 1330 ' Anillo
844                 .Invent.Object(11).Amount = 1
846                 Call EquiparInvItem(UserIndex, 1)
848                 Call EquiparInvItem(UserIndex, 2)
850                 Call EquiparInvItem(UserIndex, 3)
852                 Call EquiparInvItem(UserIndex, 4)
                
854                 .Stats.UserHechizos(1) = 25 'Paralizar
856                 .Stats.UserHechizos(2) = 26 'Inmovilizar
858                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
860                 .Stats.UserHechizos(4) = 51 'Descarga
862                 .Stats.UserHechizos(5) = 21 'Celeridad
864                 .Stats.UserHechizos(6) = 22 'Fuerza
866                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
868                 .Stats.UserHechizos(8) = 52 'Rafaga Ignea
870                 .Stats.UserHechizos(9) = 131 'Destierro
872                 .Stats.UserHechizos(10) = 132 'Oración divina
874                 .Stats.UserHechizos(11) = 133 'Plegaria
                
876             Case eClass.Paladin
878                 .Invent.NroItems = 10
880                 .Invent.Object(1).ObjIndex = 1906 'Armadura Dragón Negra
882                 .Invent.Object(1).Amount = 1 '
884                 .Invent.Object(2).ObjIndex = 1790 'Espada Saramiana
886                 .Invent.Object(2).Amount = 1 '
888                 .Invent.Object(3).ObjIndex = 1696 'Escudo Torre +1
890                 .Invent.Object(3).Amount = 1 '
892                 .Invent.Object(4).ObjIndex = 1762 'Casco legendario
894                 .Invent.Object(4).Amount = 1 '
896                 .Invent.Object(5).ObjIndex = 37 'Pocion Azul
898                 .Invent.Object(5).Amount = 10000
900                 .Invent.Object(6).ObjIndex = 37 'Pocion Azul
902                 .Invent.Object(6).Amount = 10000
904                 .Invent.Object(7).ObjIndex = 38 'Pocion Roja
906                 .Invent.Object(7).Amount = 10000
908                 .Invent.Object(8).ObjIndex = 38 'Pocion Roja
910                 .Invent.Object(8).Amount = 10000
912                 .Invent.Object(9).ObjIndex = 36 'Pocion Verde
914                 .Invent.Object(9).Amount = 10000
916                 .Invent.Object(10).ObjIndex = 39 'Pocion Amarilla
918                 .Invent.Object(10).Amount = 10000
920                 Call EquiparInvItem(UserIndex, 1)
922                 Call EquiparInvItem(UserIndex, 2)
924                 Call EquiparInvItem(UserIndex, 3)
926                 Call EquiparInvItem(UserIndex, 4)
                
928                 .Stats.UserHechizos(1) = 25 'Paralizar
930                 .Stats.UserHechizos(2) = 26 'Inmovilizar
932                 .Stats.UserHechizos(3) = 27 'Remover Paralisis
934                 .Stats.UserHechizos(4) = 51 'Descarga
936                 .Stats.UserHechizos(5) = 21 'Celeridad
938                 .Stats.UserHechizos(6) = 22 'Fuerza
940                 .Stats.UserHechizos(7) = 23 'Furia de Uhkrul
942                 .Stats.UserHechizos(8) = 100 'Golpe Iracundo
944                 .Stats.UserHechizos(9) = 101 'Heroismo
                
946             Case eClass.Hunter
948                 .Invent.NroItems = 11
950                 .Invent.Object(1).ObjIndex = 1907 'Armadura dragón verde
952                 .Invent.Object(1).Amount = 1 '
954                 .Invent.Object(2).ObjIndex = 1875 'Armadura dragón verde
956                 .Invent.Object(2).Amount = 1 '
958                 .Invent.Object(3).ObjIndex = 1717 'Escudo Gema (Cazador)
960                 .Invent.Object(3).Amount = 1 '
962                 .Invent.Object(4).ObjIndex = 1767 'Casco legendario
964                 .Invent.Object(4).Amount = 1 '
966                 .Invent.Object(5).ObjIndex = 1082 'Flecha Explosiva
968                 .Invent.Object(5).Amount = 10000 '
970                 .Invent.Object(6).ObjIndex = 38 '
972                 .Invent.Object(6).Amount = 10000
974                 .Invent.Object(7).ObjIndex = 38 '
976                 .Invent.Object(7).Amount = 10000
978                 .Invent.Object(8).ObjIndex = 36 '
980                 .Invent.Object(8).Amount = 10000
982                 .Invent.Object(9).ObjIndex = 36 '
984                 .Invent.Object(9).Amount = 10000
986                 .Invent.Object(10).ObjIndex = 39 '
988                 .Invent.Object(10).Amount = 10000
990                 .Invent.Object(11).ObjIndex = 39 '
992                 .Invent.Object(11).Amount = 10000
994                 Call EquiparInvItem(UserIndex, 1)
996                 Call EquiparInvItem(UserIndex, 2)
998                 Call EquiparInvItem(UserIndex, 3)
1000                 Call EquiparInvItem(UserIndex, 4)
1002                 .Stats.UserHechizos(1) = 152 'Paralizar
1004                 .Stats.UserHechizos(2) = 151 'Inmovilizar

1006             Case eClass.Warrior
1008                 .Invent.NroItems = 11
1010                 .Invent.Object(1).ObjIndex = 1908 'Armadura Dragón Legendaria
1012                 .Invent.Object(1).Amount = 1 '
1014                 .Invent.Object(2).ObjIndex = 1830 'Harbinger Kin
1016                 .Invent.Object(2).Amount = 1 '
1018                 .Invent.Object(3).ObjIndex = 1695 'Escudo Torre +1
1020                 .Invent.Object(3).Amount = 1 '
1022                 .Invent.Object(4).ObjIndex = 1768 'Casco legendario
1024                 .Invent.Object(4).Amount = 1 '
1026                 .Invent.Object(5).ObjIndex = 38 'Pocion Roja
1028                 .Invent.Object(5).Amount = 10000
1030                 .Invent.Object(6).ObjIndex = 38 'Pocion Roja
1032                 .Invent.Object(6).Amount = 10000
1034                 .Invent.Object(7).ObjIndex = 36 '
1036                 .Invent.Object(7).Amount = 10000
1038                 .Invent.Object(8).ObjIndex = 36 '
1040                 .Invent.Object(8).Amount = 10000
1042                 .Invent.Object(9).ObjIndex = 39 '
1044                 .Invent.Object(9).Amount = 10000
1046                 .Invent.Object(10).ObjIndex = 39 '
1048                 .Invent.Object(10).Amount = 10000
1050                 .Invent.Object(11).ObjIndex = 869 '
1052                 .Invent.Object(11).Amount = 1
1054                 Call EquiparInvItem(UserIndex, 1)
1056                 Call EquiparInvItem(UserIndex, 2)
1058                 Call EquiparInvItem(UserIndex, 3)
1060                 Call EquiparInvItem(UserIndex, 4)
1062                 Call EquiparInvItem(UserIndex, 11)
1064                 .Stats.UserHechizos(1) = 152 'Paralizar
1066                 .Stats.UserHechizos(2) = 151 'Inmovilizar

               End Select
    
1068         Call UpdateUserHechizos(True, UserIndex, 0)
        
1070         Call UpdateUserInv(True, UserIndex, 0)
        
           End With
        
         
           Exit Sub

AumentarPJ_Err:
1072      Call RegistrarError(Err.Number, Err.description, "ModBattle.AumentarPJ", Erl)
1074      Resume Next
         
End Sub

Sub RelogearUser(ByVal UserIndex As Integer, ByRef name As String, ByRef UserCuenta As String)

        On Error GoTo ErrHandler

        'Reseteamos los FLAGS
100     UserList(UserIndex).flags.Escondido = 0
102     UserList(UserIndex).flags.TargetNPC = 0
104     UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
106     UserList(UserIndex).flags.TargetObj = 0
108     UserList(UserIndex).flags.TargetUser = 0
110     UserList(UserIndex).Char.FX = 0

        'Cargamos el personaje
        Dim Leer As New clsIniReader

112     Call Leer.Initialize(CharPath & UCase$(name) & ".chr")

        'Cargamos los datos del personaje
114     Call LoadUserInit(UserIndex, Leer)

116     Call LoadUserStats(UserIndex, Leer)

118     Set Leer = Nothing

120     If UserList(UserIndex).Invent.EscudoEqpSlot = 0 Then UserList(UserIndex).Char.ShieldAnim = NingunEscudo
122     If UserList(UserIndex).Invent.CascoEqpSlot = 0 Then UserList(UserIndex).Char.CascoAnim = NingunCasco
124     If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then UserList(UserIndex).Char.WeaponAnim = NingunArma

126     Call UpdateUserInv(True, UserIndex, 0)
128     Call UpdateUserHechizos(True, UserIndex, 0)

130     If UserList(UserIndex).Correo.NoLeidos > 0 Then
132         Call WriteCorreoPicOn(UserIndex)

        End If

134     If UserList(UserIndex).flags.Paralizado Then
136         Call WriteParalizeOK(UserIndex)

        End If

138     If UserList(UserIndex).flags.Inmovilizado Then
140         Call WriteInmovilizaOK(UserIndex)

        End If

        ''
        'TODO : Feo, esto tiene que ser parche cliente
142     If UserList(UserIndex).flags.Estupidez = 0 Then
144         Call WriteDumbNoMore(UserIndex)

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
146         UserList(UserIndex).LogOnTime = Now
        #End If

148     UserList(UserIndex).Char.speeding = VelocidadNormal
150     Call WriteVelocidadToggle(UserIndex)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex))

        ''[/el oso]

        'LADDER NO SE SI QUEDO...
        'Call WriteErrorMsg(UserIndex, "LLegue 4")
152     Call WriteUpdateUserStats(UserIndex)

154     Call WriteUpdateHungerAndThirst(UserIndex)

        'Actualiza el Num de usuarios
        'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!

156     Call WriteFYA(UserIndex)

158     If UserList(UserIndex).flags.Montado = 1 Then
160         UserList(UserIndex).Char.speeding = VelocidadMontura
162         Call WriteEquiteToggle(UserIndex)
    
        End If

164     If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
166         Call WriteSafeModeOff(UserIndex)
168         UserList(UserIndex).flags.Seguro = False
        Else
170         UserList(UserIndex).flags.Seguro = True
172         Call WriteSafeModeOn(UserIndex)

        End If

        'Call modGuilds.SendGuildNews(UserIndex)

        'Load the user statistics
        'Call Statistics.UserConnected(UserIndex)

        'Call MostrarNumUsers

    

        'Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).name) & ".chr")

174     UserList(UserIndex).flags.BattleModo = 0

        Exit Sub

ErrHandler:
176     Call WriteShowMessageBox(UserIndex, "El personaje contiene un error, comuniquese con un miembro del staff.")
    

        'N = FreeFile
        'Log
        'Open App.Path & "\logs\Connect.log" For Append Shared As #N
        'Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
        'Close #N

End Sub

