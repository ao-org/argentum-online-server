Attribute VB_Name = "AIv2"
Option Explicit

Public Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
    On Error GoTo ErrorHandler
  
    Dim UserIndex As Integer
    Dim i         As Long
    Dim minDistancia As Integer

    ' Numero muy grande para que siempre haya un mínimo
    minDistancia = 32000
    
    With NpcList(NpcIndex)

        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UserIndex = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                    
            If EsObjetivoValido(NpcIndex, UserIndex) Then
            
                If Distancia(UserList(UserIndex).Pos, .Pos) < minDistancia Then
                    .Target = UserIndex
                    minDistancia = Distancia(UserList(UserIndex).Pos, .Pos)
                End If
                
            End If
             
        Next i

        If .Target > 0 Then
            If InRangoVision(.Target, .Pos.X, .Pos.Y) Then
                Call AI_AtacarObjetivo(NpcIndex)
            Else
                ' El usuario se alejo demasiado.
                .Target = 0
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
        
        ' Esta funcion espera que el target este seteado.
        If .Target = 0 Then GoTo ErrorHandler
        
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

Private Function EsObjetivoValido(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    
    ' Esto se ejecuta cuando el NPC NO tiene ningun objetivo en primer lugar.
    
    Dim RangoX    As Byte
    Dim RangoY    As Byte
        
    With NpcList(NpcIndex)
        RangoX = IIf(.Distancia <> 0, .Distancia, RANGO_VISION_X)
        RangoY = IIf(.Distancia <> 0, .Distancia, RANGO_VISION_Y)
        
    End With
    
    EsObjetivoValido = (EnRangoVision(NpcIndex, UserIndex, RangoX, RangoY) And PuedeAtacarUser(UserIndex))
    
End Function

Private Function ValidarObjetivo(NpcIndex As Integer, UserIndex As Integer) As Boolean
    
    ' Validamos al objetivo que ya estaba previamente establecido en BuscarObjetivo()
    
    With NpcList(NpcIndex)
    
        ValidarObjetivo = (.Target <> 0 And InRangoVision(UserIndex, RANGO_VISION_X, RANGO_VISION_Y) And PuedeAtacarUser(.Target))
    
    End With
    
End Function

Private Function EnRangoVision(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Limite_X As Byte, ByVal Limite_Y As Integer) As Boolean
    
    EnRangoVision = (Abs(UserList(UserIndex).Pos.X - NpcList(NpcIndex).Pos.X) <= Limite_X And Abs(UserList(UserIndex).Pos.Y - NpcList(NpcIndex).Pos.Y) <= Limite_Y)

End Function

Private Function PuedeAtacarUser(ByVal targetUserIndex As Integer) As Boolean
    
    With UserList(targetUserIndex)
            
        PuedeAtacarUser = (.flags.Muerto = 0 And .flags.invisible = 0 And .flags.Inmunidad = 0 And .flags.Oculto = 0 And .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And Not EsGM(targetUserIndex) And Not .flags.EnConsulta)
                                
    End With

End Function
