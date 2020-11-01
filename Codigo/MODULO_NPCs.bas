Attribute VB_Name = "NPCs"
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Const MaxRespawn  As Integer = 255

Public RespawnList(1 To MaxRespawn) As npc


Option Explicit

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'********************************************************
'Author: Unknown
'Llamado cuando la vida de un NPC llega a cero.
'Last Modify Date: 24/01/2007
'22/06/06: (Nacho) Chequeamos si es pretoriano
'24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
'********************************************************
On Error GoTo Errhandler
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Dim EraCriminal As Byte
    Dim TiempoRespw As Integer
    TiempoRespw = Npclist(NpcIndex).Contadores.InvervaloRespawn
   'Familiares
 '  If UserList(UserIndex).Familiar.Existe = 1 Then
      '  If UserList(UserIndex).Familiar.Invocado = 1 Then
        '   If NpcIndex = UserList(UserIndex).Familiar.Id Then
                'Call WriteConsoleMsg(UserIndex, "Tu familiar a muerto, deberas resucitarlo.", FontTypeNames.FONTTYPE_WARNING)
             '   Call WriteLocaleMsg(UserIndex, "181", FontTypeNames.FONTTYPE_WARNING)
             '   UserList(UserIndex).Familiar.Muerto = 1
           ' End If
       ' End If
   ' End If
    'Familiares
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    

    
    If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.x, MiNPC.Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("28", MiNPC.Pos.x, MiNPC.Pos.Y))
        
        End If
        
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        
       'If MiNPC.SubeSupervivencia = 1 Then
            Call SubirSkill(UserIndex, eSkill.Supervivencia)
        'End If

        
        '[KEVIN]
        If MiNPC.flags.ExpCount > 0 Then

                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount
                If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                    UserList(UserIndex).Stats.Exp = MAXEXP
               ' Call WriteConsoleMsg(UserIndex, "ID*140*" & MiNPC.flags.ExpCount, FontTypeNames.FONTTYPE_EXP)
               ' Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, MiNPC.flags.ExpCount)
               'Call WriteExpOverHead(UserIndex, MiNPC.flags.ExpCount, UserList(UserIndex).Char.CharIndex)
               Call WriteRenderValueMsg(UserIndex, MiNPC.Pos.x, MiNPC.Pos.Y - 1, MiNPC.flags.ExpCount, 6)
                
        
            MiNPC.flags.ExpCount = 0
        End If
        
        '[/KEVIN]
       ' Call WriteConsoleMsg(UserIndex, "Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, "184", FontTypeNames.FONTTYPE_FIGHT, "la criatura")
        End If
        
        'Particula al matar
       ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(MiNPC.Pos.X, MiNPC.Pos.Y, 84, 2))
        
        'Call WriteEfectOverHead(SendTarget.ToPCArea, MiNPC.GiveGLD, CStr(Npclist(NpcIndex).Char.CharIndex))
       ' Call WriteConsoleMsg(UserIndex, MiNPC.GiveGLD, FontTypeNames.FONTTYPE_FIGHT)
        
        
        If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then _
            UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
          ' Call CheckearRecompesas(UserIndex, 1)
        
        EraCriminal = Status(UserIndex)
        
        'If MiNPC.Stats.Alineacion = 0 Then
          '  If MiNPC.Numero = Guardias Then
                'UserList(UserIndex).Reputacion.NobleRep = 0

               
            'End If
        'ElseIf MiNPC.Stats.Alineacion = 1 Then
           ' UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlCAZADOR
       ' ElseIf MiNPC.Stats.Alineacion = 2 Then
            'UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlASESINO / 2
            
        'ElseIf MiNPC.Stats.Alineacion = 4 Then
           ' UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlCAZADOR
            
        'End If
       ' If Status(UserIndex) = 0 And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
       ' If Status(UserIndex) = 2 And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
        
        'If EraCriminal = 2 And Status(UserIndex) < 2 Then
        '    Call RefreshCharStatus(UserIndex)
        'ElseIf EraCriminal < 2 And Status(UserIndex) = 2 Then
        '    Call RefreshCharStatus(UserIndex)
        'End If
        
        If MiNPC.GiveEXPClan > 0 Then
            If UserList(UserIndex).GuildIndex > 0 Then
                Call modGuilds.CheckClanExp(UserIndex, MiNPC.GiveEXPClan)
          ' Else
               ' Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan, experiencia perdida.", FontTypeNames.FONTTYPE_INFOIAO)
            End If
        End If
        
        Dim i As Long, j As Long
        
            For i = 1 To MAXUSERQUESTS
        
                With UserList(UserIndex).QuestStats.Quests(i)
        
                    If .QuestIndex Then
                        If QuestList(.QuestIndex).RequiredNPCs Then
        
                            For j = 1 To QuestList(.QuestIndex).RequiredNPCs
        
                                If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
                                    If QuestList(.QuestIndex).RequiredNPC(j).Amount > .NPCsKilled(j) Then
                                        .NPCsKilled(j) = .NPCsKilled(j) + 1
        
                                    End If
                                    
                                    If QuestList(.QuestIndex).RequiredNPC(j).Amount = .NPCsKilled(j) Then
                                        Call WriteConsoleMsg(UserIndex, "Ya has matado todos los " & MiNPC.name & " que la mision " & QuestList(.QuestIndex).nombre & " requeria. Chequeá si ya estas listo para recibir la recompensa.", FontTypeNames.FONTTYPE_INFOIAO)
                                    
                                    End If
                                    
        
                                End If
        
                            Next j
        
                        End If
        
                    End If
        
                End With
        
            Next i

        
        
        Call CheckUserLevel(UserIndex)
    End If ' UserIndex > 0
   

        'Tiramos el oro
        Call NPCTirarOro(MiNPC, UserIndex)
       

        'Item Magico!
        Call NpcDropeo(MiNPC, UserIndex)
        

        
        
        'Tiramos el inventario
        Call NPC_TIRAR_ITEMS(MiNPC)
        'ReSpawn o no

        If TiempoRespw = 0 Then
            Call ReSpawnNpc(MiNPC)
        Else
            Dim indice As Integer
            MiNPC.flags.NPCActive = True
            indice = ObtenerIndiceRespawn
            RespawnList(indice) = MiNPC

        End If
        
        
    
   

    
Exit Sub

Errhandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .Attacking = 0
        .backup = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .LanzaSpells = 0
        .GolpeExacto = 0
        .invisible = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .UseAINow = False
        .AtacaAPJ = 0
        .AtacaANPC = 0
        .AIAlineacion = e_Alineacion.ninguna
        .AIPersonalidad = e_Personalidad.ninguna
    End With
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Contadores.Paralisis = 0
Npclist(NpcIndex).Contadores.TiempoExistencia = 0
Npclist(NpcIndex).Contadores.IntervaloMovimiento = 0
Npclist(NpcIndex).Contadores.IntervaloAtaque = 0
Npclist(NpcIndex).Contadores.InvervaloLanzarHechizo = 0
Npclist(NpcIndex).Contadores.InvervaloRespawn = 0



End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

Npclist(NpcIndex).Char.Body = 0
Npclist(NpcIndex).Char.CascoAnim = 0
Npclist(NpcIndex).Char.CharIndex = 0
Npclist(NpcIndex).Char.FX = 0
Npclist(NpcIndex).Char.Head = 0
Npclist(NpcIndex).Char.heading = 0
Npclist(NpcIndex).Char.loops = 0
Npclist(NpcIndex).Char.ShieldAnim = 0
Npclist(NpcIndex).Char.WeaponAnim = 0


End Sub


Sub ResetNpcCriatures(ByVal NpcIndex As Integer)


Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroCriaturas
    Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
    Npclist(NpcIndex).Criaturas(j).NpcName = vbNullString
Next j

Npclist(NpcIndex).NroCriaturas = 0

End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NroExpresiones
    Npclist(NpcIndex).Expresiones(j) = vbNullString
Next j

Npclist(NpcIndex).NroExpresiones = 0

End Sub
Sub ResetDrop(ByVal NpcIndex As Integer)

Dim j As Integer
For j = 1 To Npclist(NpcIndex).NumQuiza
    Npclist(NpcIndex).QuizaDropea(j) = 0
Next j

Npclist(NpcIndex).NumQuiza = 0

End Sub


Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    Npclist(NpcIndex).Attackable = 0
    Npclist(NpcIndex).CanAttack = 0
    Npclist(NpcIndex).Comercia = 0
    Npclist(NpcIndex).GiveEXP = 0
    Npclist(NpcIndex).GiveEXPClan = 0
    Npclist(NpcIndex).GiveGLD = 0
    Npclist(NpcIndex).Hostile = 0
    Npclist(NpcIndex).InvReSpawn = 0
    Npclist(NpcIndex).level = 0
    
    If Npclist(NpcIndex).MaestroNpc > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNpc)
    
    Npclist(NpcIndex).MaestroNpc = 0
    
    Npclist(NpcIndex).Mascotas = 0
    Npclist(NpcIndex).Movement = 0
    Npclist(NpcIndex).name = "NPC SIN INICIAR"
    Npclist(NpcIndex).NPCtype = 0
    Npclist(NpcIndex).Numero = 0
    Npclist(NpcIndex).Orig.Map = 0
    Npclist(NpcIndex).Orig.x = 0
    Npclist(NpcIndex).Orig.Y = 0
    Npclist(NpcIndex).PoderAtaque = 0
    Npclist(NpcIndex).PoderEvasion = 0
    Npclist(NpcIndex).Pos.Map = 0
    Npclist(NpcIndex).Pos.x = 0
    Npclist(NpcIndex).Pos.Y = 0
    Npclist(NpcIndex).Target = 0
    Npclist(NpcIndex).TargetNPC = 0
    Npclist(NpcIndex).TipoItems = 0
    Npclist(NpcIndex).Veneno = 0
    Npclist(NpcIndex).Desc = vbNullString
    
    
    Dim j As Integer
    For j = 1 To Npclist(NpcIndex).NroSpells
        Npclist(NpcIndex).Spells(j) = 0
    Next j
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
    Call ResetDrop(NpcIndex)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo Errhandler

    Npclist(NpcIndex).flags.NPCActive = False
    
    If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y) Then
        Call EraseNPCChar(NpcIndex)
    End If
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If

Exit Sub

Errhandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
    
    If LegalPos(Pos.Map, Pos.x, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = _
        MapData(Pos.Map, Pos.x, Pos.Y).trigger <> 3 And _
        MapData(Pos.Map, Pos.x, Pos.Y).trigger <> 2 And _
        MapData(Pos.Map, Pos.x, Pos.Y).trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim Map As Integer
Dim x As Integer
Dim Y As Integer


    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    

    
    
    If nIndex = 0 Then Exit Sub
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.Map, OrigPos.x, OrigPos.Y) Then
        
        Map = OrigPos.Map
        x = OrigPos.x
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
       
    Else
        
        Pos.Map = Mapa 'mapa
        altpos.Map = Mapa
        
        Do While Not PosicionValida
            Pos.x = RandomNumber(MinXBorder + 2, MaxXBorder - 2) 'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(MinYBorder + 2, MaxYBorder - 2) 'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            If newpos.x <> 0 And newpos.Y <> 0 Then
                altpos.x = newpos.x
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)
                If newpos.x <> 0 And newpos.Y <> 0 Then
                    altpos.x = newpos.x
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
                End If
            End If
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.Map, newpos.x, newpos.Y, PuedeAgua) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.Map = newpos.Map
                Npclist(nIndex).Pos.x = newpos.x
                Npclist(nIndex).Pos.Y = newpos.Y
                PosicionValida = True
            Else
                newpos.x = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.x <> 0 And altpos.Y <> 0 Then
                    Map = altpos.Map
                    x = altpos.x
                    Y = altpos.Y
                    Npclist(nIndex).Pos.Map = Map
                    Npclist(nIndex).Pos.x = x
                    Npclist(nIndex).Pos.Y = Y
                    Call MakeNPCChar(True, Map, nIndex, Map, x, Y)
                    Exit Sub
                Else
                    altpos.x = 50
                    altpos.Y = 50
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.x <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.Map = newpos.Map
                        Npclist(nIndex).Pos.x = newpos.x
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call MakeNPCChar(True, newpos.Map, nIndex, newpos.Map, newpos.x, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        Map = newpos.Map
        x = Npclist(nIndex).Pos.x
        Y = Npclist(nIndex).Pos.Y
    End If
    
    
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, x, Y)

End Sub

Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)
Dim CharIndex As Integer

    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(Map, x, Y).NpcIndex = NpcIndex
    
    
    Dim Simbolo As Byte
    
    Dim GG              As String

   
    GG = IIf(Npclist(NpcIndex).showName > 0, Npclist(NpcIndex).name & Npclist(NpcIndex).SubName, vbNullString)
    
    If Not toMap Then
        If Npclist(NpcIndex).QuestNumber > 0 Then
                If UserDoneQuest(sndIndex, Npclist(NpcIndex).QuestNumber) Or UserList(sndIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
                    Simbolo = 2
                Else
                    Simbolo = 1
                End If
        End If
        Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.heading, Npclist(NpcIndex).Char.CharIndex, x, Y, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 1#, True, False, 0, 0, 0, 0, Npclist(NpcIndex).Stats.MinHp, Npclist(NpcIndex).Stats.MaxHp, Simbolo)
        Call FlushBuffer(sndIndex)
    Else
        Call AgregarNpc(NpcIndex)
    End If
End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading)
    If NpcIndex > 0 Then
        Npclist(NpcIndex).Char.Body = Body
        Npclist(NpcIndex).Char.Head = Head
        Npclist(NpcIndex).Char.heading = heading
        
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, Head, heading, Npclist(NpcIndex).Char.CharIndex, 0, 0, 0, 0, 0))
    End If
End Sub

Sub EraseNPCChar(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los clientes
Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex, True))

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)

On Error GoTo errh
    Dim nPos As WorldPos
    nPos = Npclist(NpcIndex).Pos
    Call HeadtoPos(nHeading, nPos)

        ' Controlamos que la posicion sea legal, los npc que
        If LegalPosNPC(Npclist(NpcIndex).Pos.Map, nPos.x, nPos.Y, Npclist(NpcIndex).flags.AguaValida) Then
            
            If Npclist(NpcIndex).flags.AguaValida = 0 And HayAgua(Npclist(NpcIndex).Pos.Map, nPos.x, nPos.Y) Then Exit Sub
            If Npclist(NpcIndex).flags.TierraInvalida = 1 And Not HayAgua(Npclist(NpcIndex).Pos.Map, nPos.x, nPos.Y) Then Exit Sub
            
            '[Alejo-18-5]
            'server
            
            

            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(Npclist(NpcIndex).Char.CharIndex, nPos.x, nPos.Y))

            
            'Update map and user pos
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y).NpcIndex = 0
            Npclist(NpcIndex).Pos = nPos
            Npclist(NpcIndex).Char.heading = nHeading
            MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y).NpcIndex = NpcIndex
            
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        Else
            If Npclist(NpcIndex).Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                Npclist(NpcIndex).PFINFO.PathLenght = 0
            End If
        
        End If

Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)


End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo Errhandler

Dim LoopC As Integer
  
For LoopC = 1 To MAXNPCS + 1
    If LoopC > MAXNPCS Then Exit For
    If Not Npclist(LoopC).flags.NPCActive Then Exit For
Next LoopC
  
NextOpenNPC = LoopC


Exit Function
Errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer, ByVal VenenoNivel As Byte)

Dim n As Integer
n = RandomNumber(1, 100)
If n < 30 Then
    UserList(UserIndex).flags.Envenenado = VenenoNivel
   'Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, "182", FontTypeNames.FONTTYPE_FIGHT)
        End If
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean, Optional Avisar As Boolean = False) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
'***************************************************
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim Map As Integer
Dim x As Integer
Dim Y As Integer
Dim it As Integer

nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

PuedeAgua = Npclist(nIndex).flags.AguaValida
PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)

it = 0

Do While Not PosicionValida
        
        Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
        Call ClosestLegalPos(Pos, altpos, PuedeAgua)
        'Si X e Y son iguales a 0 significa que no se encontro posicion valida

        If newpos.x <> 0 And newpos.Y <> 0 Then
            'Asignamos las nuevas coordenas solo si son validas
            Npclist(nIndex).Pos.Map = newpos.Map
            Npclist(nIndex).Pos.x = newpos.x
            Npclist(nIndex).Pos.Y = newpos.Y
            PosicionValida = True
        Else
            If altpos.x <> 0 And altpos.Y <> 0 Then
                Npclist(nIndex).Pos.Map = altpos.Map
                Npclist(nIndex).Pos.x = altpos.x
                Npclist(nIndex).Pos.Y = altpos.Y
                PosicionValida = True
            Else
                PosicionValida = False
            End If
        End If
        
        it = it + 1
        
        If it > MAXSPAWNATTEMPS Then
            Call QuitarNPC(nIndex)
            SpawnNpc = 0
            Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & NpcIndex)
            Exit Function
        End If
Loop

'asignamos las nuevas coordenas
Map = newpos.Map
x = Npclist(nIndex).Pos.x
Y = Npclist(nIndex).Pos.Y

'Crea el NPC
Call MakeNPCChar(True, Map, nIndex, Map, x, Y)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, x, Y))
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
End If

If Avisar Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Npclist(nIndex).name & " ha aparecido en " & DarNameMapa(Map) & " , todo indica que puede tener una gran recompensa para el que logre sobrevivir a él.", FontTypeNames.FONTTYPE_CITIZEN))
End If


SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As npc)

If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer

Dim NpcIndex As Integer
Dim cont As Integer

'Contador
cont = 0
For NpcIndex = 1 To LastNPC

    '¿esta vivo?
    If Npclist(NpcIndex).flags.NPCActive _
       And Npclist(NpcIndex).Pos.Map = Map _
       And Npclist(NpcIndex).Hostile = 1 And _
       Npclist(NpcIndex).Stats.Alineacion = 2 Then
            cont = cont + 1
           
    End If
    
Next NpcIndex

NPCHostiles = cont

End Function

Sub NPCTirarOro(MiNPC As npc, ByVal UserIndex As Integer)

'SI EL NPC TIENE ORO LO TIRAMOS
'Pablo (ToxicWaste): Ahora se puede poner más de 10k de drop de oro en los NPC.

If MiNPC.GiveGLD > 0 Then
        If UserList(UserIndex).Grupo.EnGrupo Then
            Call CalcularDarOroGrupal(UserIndex, MiNPC.GiveGLD)
        Else
            If MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro > 99 Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro
                'Call WriteConsoleMsg(UserIndex, "¡Has ganado " & MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro & " monedas de oro!", FontTypeNames.FONTTYPE_INFOIAO)
                
               Call WriteRenderValueMsg(UserIndex, MiNPC.Pos.x, MiNPC.Pos.Y, MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro, 4)
                
                
                'Call WriteOroOverHead(UserIndex, MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro, UserList(UserIndex).Char.CharIndex)
            Else
                Dim MiObj As obj
                Dim MiAux As Double
                
                MiAux = MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro
                
                MiObj.Amount = MiAux
                MiObj.ObjIndex = iORO
                Call TirarItemAlPiso(MiNPC.Pos, MiObj)
                
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso("87", MiNPC.Pos.x, MiNPC.Pos.Y))


                Call WriteRenderValueMsg(UserIndex, MiNPC.Pos.x, MiNPC.Pos.Y, MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro, 4)
            End If
        End If
End If
End Sub

Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'conmigo. Para leer los NPCS se deberá usar la
'nueva clase clsIniReader.
'
'Alejo
'
'###################################################

Dim NpcIndex As Integer
Dim Leer As clsIniReader

Set Leer = LeerNPCs

'If requested index is invalid, abort
If Not Leer.KeyExists("NPC" & NpcNumber) Then
    OpenNPC = MAXNPCS + 1
    Exit Function
End If

NpcIndex = NextOpenNPC

If NpcIndex > MAXNPCS Then 'Limite de npcs
    OpenNPC = NpcIndex
    Exit Function
End If



Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = Leer.GetValue("NPC" & NpcNumber, "Name")
Npclist(NpcIndex).SubName = Leer.GetValue("NPC" & NpcNumber, "SubName")
Npclist(NpcIndex).Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")

Npclist(NpcIndex).Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

Npclist(NpcIndex).flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
Npclist(NpcIndex).flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
Npclist(NpcIndex).flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))

Npclist(NpcIndex).NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))


Npclist(NpcIndex).Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "Arma"))
Npclist(NpcIndex).Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "Escudo"))
Npclist(NpcIndex).Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "Casco"))

Npclist(NpcIndex).Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))


Npclist(NpcIndex).Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))

Npclist(NpcIndex).GiveEXPClan = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPClan"))

'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

Npclist(NpcIndex).Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))

Npclist(NpcIndex).flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))


Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))


Npclist(NpcIndex).QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))

Npclist(NpcIndex).PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
Npclist(NpcIndex).PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))



Npclist(NpcIndex).InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).showName = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "ShowName"))
Npclist(NpcIndex).GobernadorDe = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "GobernadorDe"))


Npclist(NpcIndex).SoundOpen = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "SoundOpen"))
Npclist(NpcIndex).SoundClose = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "SoundClose"))


Npclist(NpcIndex).IntervaloAtaque = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloAtaque"))

Npclist(NpcIndex).IntervaloMovimiento = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloMovimiento"))

Npclist(NpcIndex).InvervaloLanzarHechizo = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloLanzarHechizo"))


Npclist(NpcIndex).Contadores.InvervaloRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InvervaloRespawn"))


Npclist(NpcIndex).InformarRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InformarRespawn"))

Npclist(NpcIndex).QuizaProb = val(Leer.GetValue("NPC" & NpcNumber, "QuizaProb"))

Npclist(NpcIndex).SubeSupervivencia = val(Leer.GetValue("NPC" & NpcNumber, "SubeSupervivencia"))



If Npclist(NpcIndex).IntervaloMovimiento = 0 Then
    Npclist(NpcIndex).IntervaloMovimiento = 380
End If

If Npclist(NpcIndex).InvervaloLanzarHechizo = 0 Then
    Npclist(NpcIndex).InvervaloLanzarHechizo = 8000
End If


If Npclist(NpcIndex).IntervaloAtaque = 0 Then
    Npclist(NpcIndex).IntervaloAtaque = 2000
End If





Npclist(NpcIndex).Stats.MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
Npclist(NpcIndex).Stats.Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
    ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
    Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
    Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
Next LoopC

Npclist(NpcIndex).flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
If Npclist(NpcIndex).flags.LanzaSpells > 0 Then ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
    Npclist(NpcIndex).Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
Next LoopC


If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
    Npclist(NpcIndex).NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
    ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
    For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
        Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
        Npclist(NpcIndex).Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
    Next LoopC
End If


Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False

If Respawn Then
    Npclist(NpcIndex).flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
Else
    Npclist(NpcIndex).flags.Respawn = 1
End If

Npclist(NpcIndex).flags.backup = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
Npclist(NpcIndex).flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
Npclist(NpcIndex).flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))


Npclist(NpcIndex).flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
Npclist(NpcIndex).flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
Npclist(NpcIndex).flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Dim aux As String
aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
If LenB(aux) = 0 Then
    Npclist(NpcIndex).NroExpresiones = 0
Else
    Npclist(NpcIndex).NroExpresiones = val(aux)
    ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String
    For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
        Npclist(NpcIndex).Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
    Next LoopC
End If



'<<<<<<<<<<<<<< Sistema de Dropeo NUEVO >>>>>>>>>>>>>>>>
aux = Leer.GetValue("NPC" & NpcNumber, "NumQuiza")
If LenB(aux) = 0 Then
    Npclist(NpcIndex).NumQuiza = 0
Else
    Npclist(NpcIndex).NumQuiza = val(aux)
    ReDim Npclist(NpcIndex).QuizaDropea(1 To Npclist(NpcIndex).NumQuiza) As String
    For LoopC = 1 To Npclist(NpcIndex).NumQuiza
        Npclist(NpcIndex).QuizaDropea(LoopC) = Leer.GetValue("NPC" & NpcNumber, "QuizaDropea" & LoopC)
    Next LoopC
End If


'<<<<<<<<<<<<<< Sistema de Viajes NUEVO >>>>>>>>>>>>>>>>
aux = Leer.GetValue("NPC" & NpcNumber, "NumDestinos")
If LenB(aux) = 0 Then
    Npclist(NpcIndex).NumDestinos = 0
Else
    Npclist(NpcIndex).NumDestinos = val(aux)
    ReDim Npclist(NpcIndex).Dest(1 To Npclist(NpcIndex).NumDestinos) As String
    For LoopC = 1 To Npclist(NpcIndex).NumDestinos
        Npclist(NpcIndex).Dest(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Dest" & LoopC)
    Next LoopC
End If

'<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

Npclist(NpcIndex).Interface = val(Leer.GetValue("NPC" & NpcNumber, "Interface"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))

'Update contadores de NPCs
If NpcIndex > LastNPC Then LastNPC = NpcIndex
NumNPCs = NumNPCs + 1


'Devuelve el nuevo Indice
OpenNPC = NpcIndex

End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)

If Npclist(NpcIndex).flags.Follow Then
  Npclist(NpcIndex).flags.AttackedBy = vbNullString
  Npclist(NpcIndex).flags.Follow = False
  Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
  Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
Else
  Npclist(NpcIndex).flags.AttackedBy = UserName
  Npclist(NpcIndex).flags.Follow = True
  Npclist(NpcIndex).Movement = 4 'follow
  Npclist(NpcIndex).Hostile = 0
End If

End Sub

Public Function ObtenerIndiceRespawn() As Integer
On Error GoTo Errhandler

Dim LoopC As Integer

For LoopC = 1 To MaxRespawn
    'If LoopC > MaxRespawn Then Exit For
    If Not RespawnList(LoopC).flags.NPCActive Then Exit For
Next LoopC
  
ObtenerIndiceRespawn = LoopC


Exit Function
Errhandler:
    Call LogError("Error en ObtenerIndiceRespawn")
    
End Function

