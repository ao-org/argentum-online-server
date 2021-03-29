Attribute VB_Name = "AI"
Option Explicit

Public Enum TipoAI

    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    GuardiasAtacanCiudadanos = 6
    SigueAmo = 8
    NpcAtacaNpc = 9

    ' Animado
    Caminata = 20

    ' Eventos
    Invasion = 21

End Enum

' WyroX: Hardcodeada de la vida...
Public Const ELEMENTALFUEGO  As Integer = 962
Public Const ELEMENTALTIERRA As Integer = 961
Public Const ELEMENTALAGUA   As Integer = 960
Public Const ELEMENTALVIENTO As Integer = 963
Public Const FUEGOFATUO      As Integer = 964

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X  As Byte = 11
Public Const RANGO_VISION_Y  As Byte = 9

Public Enum e_Alineacion
    ninguna = 0
    Real = 1
    Caos = 2
    Neutro = 3
End Enum

Public Enum e_Personalidad

    ''Inerte: no tiene objetivos de ningun tipo (npcs vendedores, curas, etc)
    ''Agresivo no magico: Su objetivo es acercarse a las victimas para atacarlas
    ''Agresivo magico: Su objetivo es mantenerse lo mas lejos posible de sus victimas y atacarlas con magia
    ''Mascota: Solo ataca a quien ataque a su amo.
    ''Pacifico: No ataca.
    ninguna = 0
    Inerte = 1
    AgresivoNoMagico = 2
    AgresivoMagico = 3
    Macota = 4
    Pacifico = 5

End Enum

Public Enum e_ModoBusquedaObjetivos
    NingunoEnParticular
    FaccionarioCiudadano
    FaccionarioCriminal
End Enum

Public Sub NPCAI(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim falladesc As String

    With NpcList(NpcIndex)

        ' Ningun NPC se puede mover si esta Inmovilizado o Paralizado
        If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub

        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement

            Case TipoAI.ESTATICO
                ' Es un NPC estatico, no hace nada.

            Case TipoAI.MueveAlAzar
                falladesc = " fallo al azar"

                If .NPCtype = eNPCType.GuardiaReal Then
                    Call PerseguirUsuarioCercano(NpcIndex, e_ModoBusquedaObjetivos.FaccionarioCriminal)

                ElseIf .NPCtype = eNPCType.Guardiascaos Then
                    Call PerseguirUsuarioCercano(NpcIndex, e_ModoBusquedaObjetivos.FaccionarioCiudadano)

                Else
                        
                    ' No encontro a nadie cerca, camina unos pasos en cualquier direccion.
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    Else
                        Call AnimacionIdle(NpcIndex, True)
                    End If
                        
                End If

            Case TipoAI.NpcMaloAtacaUsersBuenos
                falladesc = " fallo NpcMaloAtacaUsersBuenos"
                Call PerseguirUsuarioCercano(NpcIndex, e_ModoBusquedaObjetivos.NingunoEnParticular)

                'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)

                'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PerseguirUsuarioCercano(NpcIndex, e_ModoBusquedaObjetivos.FaccionarioCriminal)

            Case TipoAI.GuardiasAtacanCiudadanos
                Call PerseguirUsuarioCercano(NpcIndex, e_ModoBusquedaObjetivos.FaccionarioCiudadano)

            Case TipoAI.NpcAtacaNpc
                Call AI_NpcAtacaNpc(NpcIndex)

            Case TipoAI.SigueAmo
                falladesc = " fallo SigueAmo"

                Call SeguirAmo(NpcIndex)

            Case TipoAI.Caminata
                falladesc = " fallo Caminata"

                Call HacerCaminata(NpcIndex)
                    
            Case TipoAI.Invasion
                falladesc = " fallo Invasion"
                    
                Call MovimientoInvasion(NpcIndex)

        End Select

    End With

    Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & NpcList(NpcIndex).name & " " & NpcList(NpcIndex).MaestroNPC & " mapa:" & NpcList(NpcIndex).Pos.Map & " x:" & NpcList(NpcIndex).Pos.X & " y:" & NpcList(NpcIndex).Pos.Y & " Mov:" & NpcList(NpcIndex).Movement & " TargU:" & NpcList(NpcIndex).Target & " TargN:" & NpcList(NpcIndex).TargetNPC & falladesc)

    Dim MiNPC As npc: MiNPC = NpcList(NpcIndex)
    
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)

End Sub

Public Sub PerseguirUsuarioCercano(ByVal NpcIndex As Integer, Optional ByVal TipoObjetivo As e_ModoBusquedaObjetivos)
    On Error GoTo ErrorHandler
    
    ' Buscas dentro del area de vision (donde se encuentra el NPC) el objetivo mas cercano de cierto tipo para atacar.
    
    Dim UserIndex    As Integer
    Dim i            As Long
    Dim minDistancia As Integer

    ' Numero muy grande para que siempre haya un mínimo
    minDistancia = 32000
    
    With NpcList(NpcIndex)

        ' Busco un objetivo en el area.
        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
            If EsObjetivoValido(NpcIndex, UserIndex, TipoObjetivo) Then
                    
                ' Seteo el objetivo MAS CERCANO al NPC
                If Distancia(UserList(UserIndex).Pos, .Pos) < minDistancia Then
                    .Target = UserIndex
                    minDistancia = Distancia(UserList(UserIndex).Pos, .Pos)
                End If
                    
            End If
                 
        Next i

        ' Si el NPC ya tiene un objetivo
        If .Target > 0 Then
            
            ' Vuelvo a chequear que sea valido antes de atacar
            If EsObjetivoValido(NpcIndex, UserIndex, TipoObjetivo) Then
                Call AI_AtacarObjetivo(NpcIndex)
                
            Else
                ' No encontro a nadie cerca, camina unos pasos en cualquier direccion.
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                Else
                    Call AnimacionIdle(NpcIndex, True)
                End If

                ' El usuario se alejo demasiado.
                Call RestoreOldMovement(NpcIndex)
                
            End If
       
        End If

    End With
    
    Exit Sub
    
ErrorHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "AIv2.IrUsuarioCercano", Erl)
    
End Sub

Private Sub AI_AtacarObjetivo(ByVal AtackerNpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim PegoConMagia        As Boolean
    Dim EstaLejosDelUsuario As Boolean
    Dim tHeading As Byte
    
    With NpcList(AtackerNpcIndex)
        
        If .Target = 0 Then Exit Sub
        
        EstaLejosDelUsuario = (Distancia(.Pos, UserList(.Target).Pos) > 1)
        PegoConMagia = (.flags.LanzaSpells And (RandomNumber(1, 2) = 1 Or .flags.Inmovilizado Or EstaLejosDelUsuario))

        If PegoConMagia Then
        
            ' Le lanzo un Hechizo
            Call NpcLanzaUnSpell(AtackerNpcIndex, .Target)
                
        ElseIf EstaLejosDelUsuario Then
        
            ' Camino hacia el Usuario
            tHeading = FindDirectionEAO(.Pos, UserList(.Target).Pos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
            Call MoveNPCChar(AtackerNpcIndex, tHeading)
                
        Else
            
            ' Se da vuelta y enfrenta al Usuario
            tHeading = FindDirectionEAO(.Pos, UserList(.Target).Pos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
            Call AnimacionIdle(AtackerNpcIndex, True)
            Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)
            
            ' Le pego al Usuario
            Call NpcAtacaUser(AtackerNpcIndex, .Target, tHeading)
                
        End If
        
    End With
    
    Exit Sub
    
ErrorHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "AIv2.AI_AtacarObjetivo", Erl)
    
End Sub

Public Sub AI_NpcAtacaNpc(ByVal NpcIndex As Integer)
    
    Dim tHeading As Integer
    
    With NpcList(NpcIndex)
    
        If .TargetNPC > 0 And InRangoVisionNPC(NpcIndex, NpcList(.TargetNPC).Pos.X, NpcList(.TargetNPC).Pos.Y) Then
           
            ' Me fijo si el NPC esta al lado del Objetivo
            If Distancia(.Pos, NpcList(.TargetNPC).Pos) = 1 Then
                Call SistemaCombate.NpcAtacaNpc(NpcIndex, .TargetNPC)
            End If

            ' Si NO esta inmovilizado el NPC, caminamos hacia el objetivo.
            If Not .flags.Inmovilizado Then
                
                tHeading = FindDirectionEAO(.Pos, NpcList(.TargetNPC).Pos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
                
                ' Si el NPC esta al lado del Objetivo
                If Distancia(.Pos, NpcList(.TargetNPC).Pos) = 1 Then
                    
                    ' Cambio el Heading
                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
                    Call AnimacionIdle(NpcIndex, True)

                Else
                    
                    ' Camino hacia el NPC
                    Call MoveNPCChar(NpcIndex, tHeading)
                    
                End If
                                
            End If
 
        Else
            
            Call RestoreOldMovement(NpcIndex)
            
        End If
    
    End With
    
End Sub

Public Sub SeguirAgresor(ByVal NpcIndex As Integer)
    
    ' La IA que se ejecuta cuando alguien le pega al maestro de una Mascota/Elemental
    '   o si atacas a los NPCs con Movement = 3 (TIPOAI.NPCDEFENSA)
    
    ' A diferencia de IrUsuarioCercano(), aca no buscamos objetivos cercanos en el area
    ' porque ya establecemos como objetivo a el usuario que ataco a los NPC con este tipo de IA
    
    With NpcList(NpcIndex)
        
        If EsObjetivoValido(NpcIndex, .Target, e_ModoBusquedaObjetivos.NingunoEnParticular) Then
        
            Call AI_AtacarObjetivo(NpcIndex)
        
        Else
        
            Call RestoreOldMovement(NpcIndex)
        
        End If
        
    End With
    
End Sub

Public Sub SeguirAmo(ByVal NpcIndex As Integer)
    
    Dim UserIndex As Integer
    Dim tHeading As Integer
    
    With NpcList(NpcIndex)
        
        If .MaestroUser = 0 Then Exit Sub
        
        ' Si la mascota no tiene objetivo establecido.
        If .Target = 0 And .TargetNPC = 0 Then
            
            UserIndex = .MaestroUser
            
            If EnRangoVision(NpcIndex, .MaestroUser, RANGO_VISION_X, RANGO_VISION_Y) Then
                
                If UserList(UserIndex).flags.Muerto = 0 And _
                    UserList(UserIndex).flags.invisible = 0 And _
                    UserList(UserIndex).flags.Oculto = 0 And _
                    Distancia(.Pos, UserList(UserIndex).Pos) > 3 Then
                    
                    ' Caminamos cerca del usuario
                    tHeading = FindDirectionEAO(.Pos, UserList(UserIndex).Pos)
                    Call MoveNPCChar(NpcIndex, tHeading)
                    Exit Sub
                    
                Else
                    
                    ' Caminamos aleatoriamente por ahi cerca.
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    Else
                        Call AnimacionIdle(NpcIndex, True)

                    End If
                
                End If
                
            End If
            
        End If
        
    End With
    
    Call RestoreOldMovement(NpcIndex)
    
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

    With NpcList(NpcIndex)
        
        .Target = 0
        
        ' Si el NPC no tiene maestro, reseteamos el movimiento que tenia antes.
        If .MaestroUser = 0 Then
        
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString

        Else
            
            ' Si tiene maestro, hacemos que lo siga.
            Call FollowAmo(NpcIndex)
            
        End If

    End With

End Sub

Private Sub HacerCaminata(ByVal NpcIndex As Integer)
    On Error GoTo Handler
    
    Dim Destino As WorldPos
    Dim Heading As eHeading
    Dim NextTile As WorldPos
    Dim MoveChar As Integer
    Dim PudoMover As Boolean

    With NpcList(NpcIndex)
    
        Destino.Map = .Pos.Map
        Destino.X = .Orig.X + .Caminata(.CaminataActual).Offset.X
        Destino.Y = .Orig.Y + .Caminata(.CaminataActual).Offset.Y

        ' Si todavía no llegó al destino
        If .Pos.X <> Destino.X Or .Pos.Y <> Destino.Y Then
        
            ' Tratamos de acercarnos (podemos pisar npcs, usuarios o triggers)
            Heading = FindDirectionEAO(.Pos, Destino, .flags.AguaValida, .flags.TierraInvalida = 0, True, True)
            
            ' Obtengo la posición según el heading
            NextTile = .Pos
            Call HeadtoPos(Heading, NextTile)
            
            ' Si hay un NPC
            MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).NpcIndex
            If MoveChar Then
                ' Lo movemos hacia un lado
                Call MoveNpcToSide(MoveChar, Heading)
            End If
            
            ' Si hay un user
            MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).UserIndex
            If MoveChar Then
                ' Si no está muerto o es admin invisible (porque a esos los atraviesa)
                If UserList(MoveChar).flags.AdminInvisible = 0 Or UserList(MoveChar).flags.Muerto = 0 Then
                    ' Lo movemos hacia un lado
                    Call MoveUserToSide(MoveChar, Heading)
                End If
            End If
            
            ' Movemos al NPC de la caminata
            PudoMover = MoveNPCChar(NpcIndex, Heading)
            
            ' Si no pudimos moverlo, hacemos como si hubiese llegado a destino... para evitar que se quede atascado
            If Not PudoMover Or Distancia(.Pos, Destino) = 0 Then
            
                ' Llegamos a destino, ahora esperamos el tiempo necesario para continuar
                .Contadores.IntervaloMovimiento = GetTickCount + .Caminata(.CaminataActual).Espera - .IntervaloMovimiento
                
                ' Pasamos a la siguiente caminata
                .CaminataActual = .CaminataActual + 1
                
                ' Si pasamos el último, volvemos al primero
                If .CaminataActual > UBound(.Caminata) Then
                    .CaminataActual = 1
                End If
                
            End If
            
        ' Si por alguna razón estamos en el destino, seguimos con la siguiente caminata
        Else
        
            .CaminataActual = .CaminataActual + 1
            
            ' Si pasamos el último, volvemos al primero
            If .CaminataActual > UBound(.Caminata) Then
                .CaminataActual = 1
            End If
            
        End If
    
    End With
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "AI.HacerCaminata", Erl)
    Resume Next
End Sub

Private Sub MovimientoInvasion(ByVal NpcIndex As Integer)
    On Error GoTo Handler
    
    With NpcList(NpcIndex)
        Dim SpawnBox As tSpawnBox
        SpawnBox = Invasiones(.flags.InvasionIndex).SpawnBoxes(.flags.SpawnBox)
    
        ' Calculamos la distancia a la muralla y generamos una posición de destino
        Dim DistanciaMuralla As Integer, Destino As WorldPos
        Destino = .Pos
        
        If SpawnBox.Heading = eHeading.EAST Or SpawnBox.Heading = eHeading.WEST Then
            DistanciaMuralla = Abs(.Pos.X - SpawnBox.CoordMuralla)
            Destino.X = SpawnBox.CoordMuralla
        Else
            DistanciaMuralla = Abs(.Pos.Y - SpawnBox.CoordMuralla)
            Destino.Y = SpawnBox.CoordMuralla
        End If

        ' Si todavía está lejos de la muralla
        If DistanciaMuralla > 1 Then
        
            ' Tratamos de acercarnos (sin pisar)
            Dim Heading As eHeading
            Heading = FindDirectionEAO(.Pos, Destino, .flags.AguaValida, .flags.TierraInvalida = 0, True)
            
            ' Nos aseguramos que la posición nueva está dentro del rectángulo válido
            Dim NextTile As WorldPos
            NextTile = .Pos
            Call HeadtoPos(Heading, NextTile)
            
            ' Si la posición nueva queda fuera del rectángulo válido
            If Not InsideRectangle(SpawnBox.LegalBox, NextTile.X, NextTile.Y) Then
                ' Invertimos la dirección de movimiento
                Heading = InvertHeading(Heading)
            End If
            
            ' Movemos el NPC
            Call MoveNPCChar(NpcIndex, Heading)
        
        ' Si está pegado a la muralla
        Else
        
            ' Chequeamos el intervalo de ataque
            If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
                Exit Sub
            End If
            
            ' Nos aseguramos que mire hacia la muralla
            If .Char.Heading <> SpawnBox.Heading Then
                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, SpawnBox.Heading)
            End If
            
            ' Sonido de ataque (si tiene)
            If .flags.Snd1 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
            End If
            
            ' Sonido de impacto
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            ' Dañamos la muralla
            Call HacerDañoMuralla(.flags.InvasionIndex, RandomNumber(.Stats.MinHIT, .Stats.MaxHit))  ' TODO: Defensa de la muralla? No hace falta creo...

        End If
    
    End With

    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "AI.MovimientoInvasion", Erl)
    Resume Next
End Sub

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
    On Error GoTo NpcLanzaUnSpell_Err
        
    With UserList(UserIndex)
        
        If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub
        If NpcList(NpcIndex).Pos.Map <> .Pos.Map Then Exit Sub

        If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Or .flags.NoMagiaEfeceto = 1 Or .flags.EnConsulta Then Exit Sub
    
        Dim K As Integer
            K = RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells)

        Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, NpcList(NpcIndex).Spells(K))

        If NpcList(NpcIndex).Target = 0 Then NpcList(NpcIndex).Target = UserIndex

        If .flags.AtacadoPorNpc = 0 And .flags.AtacadoPorUser = 0 Then
            .flags.AtacadoPorNpc = NpcIndex
        End If
        
    End With

    Exit Sub

NpcLanzaUnSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "AI.NpcLanzaUnSpell", Erl)

    Resume Next
        
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
        
    On Error GoTo NpcLanzaUnSpellSobreNpc_Err
    
    With NpcList(NpcIndex)
        
        If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub
        If .Pos.Map <> NpcList(TargetNPC).Pos.Map Then Exit Sub
    
        Dim K As Integer
            K = RandomNumber(1, .flags.LanzaSpells)

        Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, .Spells(K))
    
    End With
     
    Exit Sub

NpcLanzaUnSpellSobreNpc_Err:
    Call RegistrarError(Err.Number, Err.Description, "AI.NpcLanzaUnSpellSobreNpc", Erl)
    Resume Next
        
End Sub




' ---------------------------------------------------------------------------------------------------
'                                       HELPERS
' ---------------------------------------------------------------------------------------------------

Private Function EsObjetivoValido(ByVal NpcIndex As Integer, _
                                  ByVal UserIndex As Integer, _
                                  ByVal ModoBusqueda As e_ModoBusquedaObjetivos) As Boolean
    
    ' Esto se ejecuta cuando el NPC NO tiene ningun objetivo en primer lugar.
    
    Dim RangoX    As Byte
    Dim RangoY    As Byte
        
    With NpcList(NpcIndex)
        
        RangoX = IIf(.Distancia <> 0, .Distancia, RANGO_VISION_X)
        RangoY = IIf(.Distancia <> 0, .Distancia, RANGO_VISION_Y)
        
    End With
    
    If UserIndex > 0 Then
        
        ' Esta condicion debe ejecutarse independiemente de el modo de busqueda.
        EsObjetivoValido = (EnRangoVision(NpcIndex, UserIndex, RangoX, RangoY) And PuedeAtacarUser(UserIndex))
        
        ' Aca tenemos ciertos criterios que podemos usar a la hora de establecer el objetivo de un NPC
        Select Case ModoBusqueda
                           
            ' Si queres buscar Criminales cercanos...
            Case e_ModoBusquedaObjetivos.FaccionarioCiudadano
                
                EsObjetivoValido = (EsObjetivoValido And (Status(UserIndex) = 1 Or Status(UserIndex) = 3))
            
            ' Si queres buscar Ciudadanos cercanos...
            Case e_ModoBusquedaObjetivos.FaccionarioCriminal
                
                EsObjetivoValido = (EsObjetivoValido And (Status(UserIndex) = 0 Or Status(UserIndex) = 2))
            
            Case Else
                ' Ok. No hay nada especial para hacer, cualquiera puede ser objetivo!
            
                            
        End Select

    Else
        
        EsObjetivoValido = False
    
    End If
    
End Function

Private Function EnRangoVision(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Limite_X As Byte, ByVal Limite_Y As Integer) As Boolean
    
    EnRangoVision = (Abs(UserList(UserIndex).Pos.X - NpcList(NpcIndex).Pos.X) <= Limite_X And Abs(UserList(UserIndex).Pos.Y - NpcList(NpcIndex).Pos.Y) <= Limite_Y)

End Function

Private Function PuedeAtacarUser(ByVal targetUserIndex As Integer) As Boolean
    
    With UserList(targetUserIndex)
            
        PuedeAtacarUser = (.flags.Muerto = 0 And .flags.invisible = 0 And .flags.Inmunidad = 0 And .flags.Oculto = 0 And .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And Not EsGM(targetUserIndex) And Not .flags.EnConsulta)
                                
    End With

End Function
