Attribute VB_Name = "ModInvasion"
Option Explicit

Type tSpawnBox
    TopLeft As WorldPos
    BottomRight As WorldPos
    Heading As eHeading
    CoordMuralla As Integer
    LegalBox As Rectangle
End Type

Type tTopInvasion
    UserName As String
    Score As Long
End Type

Type tInvasion
    Activa As Boolean
    ' Muralla
    VidaMuralla As Long
    MaxVidaMuralla As Long
    ' Users
    Top10Users(1 To 10) As tTopInvasion
    ' NPCs
    NPCsVivos() As Integer
    CantNPCs As Integer
    NumNPCsSpawn() As Integer
    MaxNPCs As Integer
    ' Aviso
    aviso As String
    AvisarTiempo As Integer
    RepetirAviso As Integer
    TimerRepetirAviso As Integer
    ' Descripción
    Desc As String
    RepetirDesc As Integer
    TimerRepetirDesc As Integer
    ' Spawns
    SpawnBoxes() As tSpawnBox
    IntervaloSpawn As Integer
    TimerSpawn As Integer
    ' Duracion e intervalos
    TimerInvasion As Integer
    Intervalo As Integer
    Duracion As Integer
    ' Mostrar info en pantalla
    TimerMostrarInfo As Integer
    TiempoDeInicio As Long
    ' Mensajes de fin
    MensajeGanaron As String
    MensajePerdieron As String
End Type

Public Invasiones() As tInvasion

Sub CargarInfoEventos()
    Dim File As clsIniReader
    Set File = New clsIniReader

    Call File.Initialize(DatPath & "Eventos.dat")
    
    Dim CantInvasiones As Integer
    CantInvasiones = val(File.GetValue("Invasiones", "Cantidad"))

    If CantInvasiones <= 0 Then Exit Sub

    ReDim Invasiones(1 To CantInvasiones)
    
    Dim i As Integer, j As Integer, nombre As String, tmpStr As String, Fields() As String
    
    For i = 1 To CantInvasiones
        
        nombre = File.GetValue("Invasiones", "Invasion" & i)
        
        With Invasiones(i)
    
            .MaxVidaMuralla = val(File.GetValue(nombre, "MaxVidaMuralla"))
            .MaxNPCs = val(File.GetValue(nombre, "MaxNPCs"))
            .aviso = File.GetValue(nombre, "Aviso")
            .AvisarTiempo = val(File.GetValue(nombre, "AvisarTiempo"))
            .RepetirAviso = val(File.GetValue(nombre, "RepetirAviso"))
            .Desc = File.GetValue(nombre, "Desc")
            .RepetirDesc = val(File.GetValue(nombre, "RepetirDesc"))
            .IntervaloSpawn = val(File.GetValue(nombre, "IntervaloSpawn"))
            .Duracion = val(File.GetValue(nombre, "Duracion"))
            .Intervalo = val(File.GetValue(nombre, "Intervalo"))
            .TimerInvasion = val(File.GetValue(nombre, "Offset"))
            .MensajeGanaron = File.GetValue(nombre, "MensajeGanaron")
            .MensajePerdieron = File.GetValue(nombre, "MensajePerdieron")
            
            If .MaxNPCs <= 0 Then Exit Sub

            ReDim .NPCsVivos(1 To .MaxNPCs)
            
            tmpStr = File.GetValue(nombre, "NPCs")
            If LenB(tmpStr) > 0 Then
                Fields = Split(tmpStr, "-")
                
                ReDim .NumNPCsSpawn(1 To UBound(Fields) + 1)
                
                For j = 1 To UBound(.NumNPCsSpawn)
                    .NumNPCsSpawn(j) = val(Fields(j - 1))
                Next
            Else
                ReDim .NumNPCsSpawn(0)
            End If
            
            Dim SpawnBoxes As Integer
            SpawnBoxes = val(File.GetValue(nombre, "SpawnBoxes"))
            
            If SpawnBoxes <= 0 Then Exit Sub
            
            ReDim .SpawnBoxes(1 To SpawnBoxes)
            
            For j = 1 To SpawnBoxes
                tmpStr = File.GetValue(nombre, "SpawnBox" & j)
                
                If LenB(tmpStr) > 0 Then
                    Fields = Split(tmpStr, "-", 7)
                
                    If UBound(Fields) = 6 Then
                        With .SpawnBoxes(j)
                            ' Mapa
                            .TopLeft.Map = val(Fields(0))
                            .BottomRight.Map = .TopLeft.Map
                            ' TopLeft
                            .TopLeft.X = val(Fields(1))
                            .TopLeft.Y = val(Fields(2))
                            ' BottomRight
                            .BottomRight.X = val(Fields(3))
                            .BottomRight.Y = val(Fields(4))
                            ' Dirección de ataque
                            .Heading = String2Heading(Fields(5))
                            .CoordMuralla = val(Fields(6))
                            ' Calculamos las posiciones válidas de los NPCs
                            .LegalBox.X1 = .TopLeft.X
                            .LegalBox.Y1 = .TopLeft.Y
                            .LegalBox.X2 = .BottomRight.X
                            .LegalBox.Y2 = .BottomRight.Y

                            Select Case .Heading
                                Case eHeading.NORTH: .LegalBox.Y1 = .CoordMuralla
                                Case eHeading.SOUTH: .LegalBox.Y2 = .CoordMuralla
                                Case eHeading.EAST: .LegalBox.X2 = .CoordMuralla
                                Case eHeading.WEST: .LegalBox.X1 = .CoordMuralla
                            End Select
                        End With
                    End If
                End If
                
            Next
            
        End With
    Next
    
    Set File = Nothing
End Sub

Sub IniciarInvasion(ByVal index As Integer)
    
    With Invasiones(index)
    
        .Activa = True
        
        .VidaMuralla = .MaxVidaMuralla
        
        .TiempoDeInicio = GetTickCount

        ' Enviamos info sobre la invasión a los usuarios en estos mapas
        Call EnviarInfoInvasion(index)
        
        Call MensajeGlobal(.Desc, FontTypeNames.FONTTYPE_New_Eventos)
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(150, NO_3D_SOUND, NO_3D_SOUND))
    
    End With
    
End Sub

Sub FinalizarInvasion(ByVal index As Integer)

    With Invasiones(index)
    
        Dim Ganaron As Boolean
        
        If .VidaMuralla > 0 Then
            Call MensajeGlobal(.MensajeGanaron, FontTypeNames.FONTTYPE_New_Eventos)
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            Ganaron = True
        Else
            Call MensajeGlobal(.MensajePerdieron, FontTypeNames.FONTTYPE_New_Eventos)
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If

        ' Limpiamos flags
        .Activa = False
        .TimerRepetirAviso = 0
        .TimerInvasion = 0
        .TimerMostrarInfo = 0
        .TimerRepetirDesc = 0
        .TimerSpawn = 0

        ' Matamos los NPCs que quedaron
        Dim i As Integer
        For i = 1 To UBound(.NPCsVivos)
            If .NPCsVivos(i) Then
                Call QuitarNPC(.NPCsVivos(i))
                .NPCsVivos(i) = 0
                
                .CantNPCs = .CantNPCs - 1
                If .CantNPCs <= 0 Then Exit For
            End If
        Next
        
        ' Entregamos premios y limpiamos el top
        Dim UserIndex As Integer, OroGanado As Long, PremioStr As String
    
        OroGanado = 50000 * OroMult
        PremioStr = "¡La ciudad te entrega " & PonerPuntos(OroGanado) & " monedas de oro por tu ayuda durante la invasión!"
        
        For i = 1 To UBound(.Top10Users)
            With .Top10Users(i)
        
                If LenB(.UserName) Then
                    If Ganaron And .Score > 0 Then
                        ' Si está conectado
                        UserIndex = NameIndex(.UserName)
                        If UserIndex Then
                            ' Le damos el oro
                            Call WriteConsoleMsg(UserIndex, PremioStr, FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + OroGanado
                            Call WriteUpdateGold(UserIndex)
                        End If
                    End If
                    
                    .UserName = vbNullString
                    .Score = 0
                End If
            End With
        Next

        ' Sacamos el cartel de la pantalla de todos
        Dim Mapa As Integer, j As Integer
        For i = 1 To UBound(.SpawnBoxes)
            Mapa = .SpawnBoxes(i).TopLeft.Map
            
            For j = 1 To ModAreas.ConnGroups(Mapa).CountEntrys
                Call WriteInvasionInfo(ModAreas.ConnGroups(Mapa).UserEntrys(j), 0, 0, 0)
            Next
        Next

    End With

End Sub

Sub InvasionSpawnNPC(ByVal index As Integer)

    With Invasiones(index)
    
        ' Si ya hay el máximo de NPCs, no spawneamos nada
        If .CantNPCs >= .MaxNPCs Then Exit Sub
        
        ' Elegimos al azar el NPC a spawnear
        Dim NpcNumber As Integer
        NpcNumber = .NumNPCsSpawn(RandomNumber(1, UBound(.NumNPCsSpawn)))
        
        ' Elegimos un área al azar (TODO: elegir con más probabilidad según más área cubra)
        Dim SpawnBox As Integer
        SpawnBox = RandomNumber(1, UBound(.SpawnBoxes))
        
        With .SpawnBoxes(SpawnBox)
        
            ' Elegimos un tile al azar dentro del área
            Dim SpawnPos As WorldPos
            SpawnPos.Map = .TopLeft.Map
            SpawnPos.X = RandomNumber(.TopLeft.X, .BottomRight.X)
            SpawnPos.Y = RandomNumber(.TopLeft.Y, .BottomRight.Y)
            
            ' Obtenemos la dirección y coordenada en la que se encuentra la muralla
            Dim Heading As eHeading
            Heading = .Heading
        
        End With

        ' Buscamos un índice vacío en el array de NPCs
        Dim i As Integer
        For i = 1 To UBound(.NPCsVivos)
            If .NPCsVivos(i) = 0 Then Exit For
        Next

        ' Spawneamos el NPC
        .NPCsVivos(i) = SpawnNpc(NpcNumber, SpawnPos, True, False)
        
        Debug.Assert .NPCsVivos(i) <> 0

        ' Si pudimos spawnearlo
        If .NPCsVivos(i) Then
            .CantNPCs = .CantNPCs + 1
    
            ' Lo colocamos mirando en dirección a la muralla
            Call ChangeNPCChar(.NPCsVivos(i), NpcList(.NPCsVivos(i)).Char.Body, NpcList(.NPCsVivos(i)).Char.Head, Heading)
            
            ' Guardamos información sobre el spawn
            NpcList(.NPCsVivos(i)).flags.InvasionIndex = index
            NpcList(.NPCsVivos(i)).flags.SpawnBox = SpawnBox
            NpcList(.NPCsVivos(i)).flags.IndexInInvasion = i
        End If

    End With

End Sub

Public Sub MuereNpcInvasion(ByVal index As Integer, ByVal NpcIndex As Integer)

    With Invasiones(index)
    
        .NPCsVivos(NpcIndex) = 0
    
        .CantNPCs = .CantNPCs - 1

    End With

End Sub

Private Function String2Heading(str As String) As eHeading

    Select Case LCase$(str)
        Case "norte": String2Heading = eHeading.NORTH
        Case "sur": String2Heading = eHeading.SOUTH
        Case "este": String2Heading = eHeading.EAST
        Case "oeste": String2Heading = eHeading.WEST
    End Select

End Function

Public Sub EnviarInfoInvasion(ByVal index As Integer)

    With Invasiones(index)
    
        Dim PorcentajeVida As Byte, PorcentajeTiempo As Byte
        
        PorcentajeVida = (.VidaMuralla / .MaxVidaMuralla) * 100
        
        PorcentajeTiempo = (GetTickCount - .TiempoDeInicio) / (.Duracion * 600)
    
        Dim i As Integer, Mapa As Integer, j As Integer
        For i = 1 To UBound(.SpawnBoxes)
            Mapa = .SpawnBoxes(i).TopLeft.Map
            
            For j = 1 To ModAreas.ConnGroups(Mapa).CountEntrys
                Call WriteInvasionInfo(ModAreas.ConnGroups(Mapa).UserEntrys(j), index, PorcentajeVida, PorcentajeTiempo)
            Next
        Next
        
    End With

End Sub

Public Sub HacerDañoMuralla(ByVal index As Integer, ByVal Daño As Long)
    
    With Invasiones(index)
    
        .VidaMuralla = .VidaMuralla - Daño
        
        If .VidaMuralla <= 0 Then
            Call FinalizarInvasion(index)
        End If
    
    End With
    
End Sub

Public Sub SumarScoreInvasion(ByVal index As Integer, ByVal UserIndex As Integer, ByVal Score As Long)
    
    With Invasiones(index)
    
        Dim i As Integer
        Dim tmpUser As tTopInvasion
        
        ' Buscamos si estamos en el top
        For i = 1 To UBound(.Top10Users)
            If LenB(.Top10Users(i).UserName) = 0 Then
                ' Llegamos a un lugar vacío, entonces no está en el top
                Exit For
            
            ElseIf .Top10Users(i).UserName = UserList(UserIndex).name Then
                ' Está en el top, así que le sumamos el puntaje
                .Top10Users(i).Score = .Top10Users(i).Score + Score
                
                ' Revisamos si subió en el top
                Dim j As Integer
                For j = i - 1 To 1 Step -1
                    ' Si el que está arriba tiene un puntaje menor, los cambiamos
                    If .Top10Users(j).Score < .Top10Users(j + 1).Score Then
                        tmpUser = .Top10Users(j)
                        .Top10Users(j) = .Top10Users(j + 1)
                        .Top10Users(j + 1) = tmpUser
                    Else
                        ' Sino, salimos
                        Exit For
                    End If
                Next
                
                ' Salimos, no hace falta agregarlo
                Exit Sub
            
            End If
        Next
        
        ' Si llegamos acá, entonces hay que meterlo al top
        For i = UBound(.Top10Users) To 1 Step -1
            ' Buscamos el lugar indicado
            If .Top10Users(i).Score > Score Then
                Exit For
            End If
        Next
        
        ' Si entró en el top
        If i < UBound(.Top10Users) Then
            ' Movemos a los que le siguen
            For j = UBound(.Top10Users) To i + 2
                .Top10Users(j) = .Top10Users(j - 1)
            Next
            
            ' Lo colocamos en la posición que le corresponde
            With .Top10Users(i + 1)
                .UserName = UserList(UserIndex).name
                .Score = Score
            End With
        End If
    
    End With
    
End Sub
