Attribute VB_Name = "AIv2"
Option Explicit

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

Public Enum TipoAI

    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    GuardiasAtacanCiudadanos = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10

    'Pretorianos
    SacerdotePretorianoAi = 11
    GuerreroPretorianoAi = 12
    MagoPretorianoAi = 13
    CazadorPretorianoAi = 14
    ReyPretoriano = 15

    ' Animado
    Caminata = 20
    
    ' Eventos
    Invasion = 21

End Enum

Public Function IrUsuarioCercano2(ByVal NpcIndex As Integer)
    
    Dim tHeading  As Byte
    Dim UserIndex As Integer
    Dim Pos       As WorldPos
    Dim i         As Long
    Dim ComoAtaco As Byte
    
    With NpcList(NpcIndex)
            
        ' Si el NPC no tiene un objetivo definido
        If .Target = 0 Then
                
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                        
                If EsObjetivoValido(NpcIndex, UserIndex) Then
                    
                    .Target = UserIndex
                    
                    Exit For
                    
                End If
                 
            Next i
      
        End If
        
        If .Target > 0 Then
            
            Call AI_AtacarObjetivo(NpcIndex)
       
        End If

    End With
    
End Function

Private Function AI_AtacarObjetivo(AtackerNpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim PegoConMagia        As Boolean
    Dim EstaLejosDelUsuario As Boolean
    
    With NpcList(NpcIndex)
        
        ' Esta funcion espera que el target este seteado.
        If .Target = 0 Then GoTo ErrorHandler
        
        EstaLejosDelUsuario = (Distancia(.Pos, UserList(.Target).Pos) > 1)
        PegoConMagia = (.flags.LanzaSpells And (RandomNumber(1, 2) = 1 Or .flags.Inmovilizado Or EstaLejosDelUsuario))

        If PegoConMagia Then
        
            ' Le lanzo un Hechizo
            Call NpcLanzaUnSpell(NpcIndex, .Target)
                
        ElseIf EstaLejosDelUsuario Then
        
            ' Camino hacia el Usuario
            tHeading = FindDirectionEAO(.Pos, UserList(.Target).Pos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
            Call MoveNPCChar(NpcIndex, tHeading)
                
        Else
            
            ' Se da vuelta y enfrenta al Usuario
            tHeading = FindDirectionEAO(.Pos, UserList(.Target).Pos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
            Call AnimacionIdle(NpcIndex, True)
            Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
            
            ' Le pego al Usuario
            Call NpcAtacaUser(NpcIndex, .Target, tHeading)
                
        End If
        
    End With
    
    Exit Function
    
ErrorHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "AIv2.AI_AtacarObjetivo", Erl)
    
End Function

Private Function EsObjetivoValido(NpcIndex As Integer, UserIndex As Integer) As Boolean
    
    ' Esto se ejecuta cuando el NPC NO tiene ningun objetivo en primer lugar.
    
    Dim UserIndex As Integer
    Dim i         As Long
    Dim RangoX    As Byte
    Dim RangoY    As Byte
    
    EsObjetivoValido = False
    
    If .Distancia <> 0 Then
        RangoX = .Distancia
        RangoY = .Distancia

    Else
        RangoX = RANGO_VISION_X
        RangoY = RANGO_VISION_Y
            
    End If

    EsObjetivoValido = (EnRangoVision(UserIndex, RangoX, RangoY) And PuedeAtacarUser(UserIndex))

End Function

Private Function ValidarObjetivo(NpcIndex As Integer, UserIndex As Integer)
    
    ' Validamos al objetivo que ya estaba previamente establecido en BuscarObjetivo()
    
    With NpcList(NpcIndex)
    
        ValidarObjetivo = (.Target <> 0 And InRangoVision(UserIndex, RANGO_VISION_X, RANGO_VISION_Y) And PuedeAtacarUser(.Target))
    
    End With
    
End Function

Private Function EnRangoVision(Index As Integer, Limite_X As Byte, Limite_Y As Integer) As Boolean
    
    EnRangoVision = (Abs(UserList(Index).Pos.X - .Pos.X) <= Limite_X And Abs(UserList(Index).Pos.Y - .Pos.Y) <= Limite_Y)
   
End Function

Private Function PuedeAtacarUser(ByVal targetUserIndex As Integer) As Boolean
    
    With UserList(targetUserIndex)
            
        PuedeAtacarUser = (.flags.Muerto = 0 And .flags.invisible = 0 And .flags.Inmunidad = 0 And .flags.Oculto = 0 And .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And Not EsGM(targetUserIndex) And Not .flags.EnConsulta)
                                
    End With

End Function
