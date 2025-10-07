Attribute VB_Name = "ModInvasion"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Type t_SpawnBox
    TopLeft As t_WorldPos
    BottomRight As t_WorldPos
    Heading As e_Heading
    CoordMuralla As Integer
    LegalBox As t_Rectangle
End Type

Type t_TopInvasion
    username As String
    Score As Long
End Type

Type t_Invasion
    Activa As Boolean
    ' Muralla
    VidaMuralla As Long
    MaxVidaMuralla As Long
    ' Users
    Top10Users(1 To 10) As t_TopInvasion
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
    SpawnBoxes() As t_SpawnBox
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

Public Invasiones() As t_Invasion

Sub CargarInfoEventos()
    Dim File As clsIniManager
    Set File = New clsIniManager
    Call File.Initialize(DatPath & "Eventos.dat")
    Dim CantInvasiones As Integer
    CantInvasiones = val(File.GetValue("Invasiones", "Cantidad"))
    If CantInvasiones <= 0 Then
        ReDim Invasiones(0)
        frmMain.Invasion.Enabled = False
        Exit Sub
    End If
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
                            .TopLeft.x = val(Fields(1))
                            .TopLeft.y = val(Fields(2))
                            ' BottomRight
                            .BottomRight.x = val(Fields(3))
                            .BottomRight.y = val(Fields(4))
                            ' Dirección de ataque
                            .Heading = String2Heading(Fields(5))
                            .CoordMuralla = val(Fields(6))
                            ' Calculamos las posiciones válidas de los NPCs
                            .LegalBox.X1 = .TopLeft.x
                            .LegalBox.Y1 = .TopLeft.y
                            .LegalBox.X2 = .BottomRight.x
                            .LegalBox.Y2 = .BottomRight.y
                            Select Case .Heading
                                Case e_Heading.NORTH: .LegalBox.Y1 = .CoordMuralla
                                Case e_Heading.SOUTH: .LegalBox.Y2 = .CoordMuralla
                                Case e_Heading.EAST: .LegalBox.X2 = .CoordMuralla
                                Case e_Heading.WEST: .LegalBox.X1 = .CoordMuralla
                            End Select
                        End With
                    End If
                End If
            Next
        End With
    Next
    frmMain.Invasion.Enabled = True
    Set File = Nothing
End Sub

Sub IniciarInvasion(ByVal Index As Integer)
    If UBound(Invasiones) = 0 Then Exit Sub
    With Invasiones(Index)
        .Activa = True
        .VidaMuralla = .MaxVidaMuralla
        .TiempoDeInicio = GetTickCountRaw()
        ' Enviamos info sobre la invasión a los usuarios en estos mapas
        Call EnviarInfoInvasion(Index)
        Call MensajeGlobal(.Desc, e_FontTypeNames.FONTTYPE_New_Eventos)
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(150, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

Sub FinalizarInvasion(ByVal Index As Integer)
    With Invasiones(Index)
        Dim Ganaron As Boolean
        If .VidaMuralla > 0 Then
            Call MensajeGlobal(.MensajeGanaron, e_FontTypeNames.FONTTYPE_New_Eventos)
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            Ganaron = True
        Else
            Call MensajeGlobal(.MensajePerdieron, e_FontTypeNames.FONTTYPE_New_Eventos)
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
                Call QuitarNPC(.NPCsVivos(i), eClearInvasion)
                .NPCsVivos(i) = 0
                .CantNPCs = .CantNPCs - 1
                If .CantNPCs <= 0 Then Exit For
            End If
        Next
        ' Entregamos premios y limpiamos el top
        Dim tUser As t_UserReference, OroGanado As Long, PremioStr As String
        OroGanado = 50000 * SvrConfig.GetValue("GoldMult")
        PremioStr = "¡La ciudad te entrega " & PonerPuntos(OroGanado) & " monedas de oro por tu ayuda durante la invasión!"
        For i = 1 To UBound(.Top10Users)
            With .Top10Users(i)
                If LenB(.username) Then
                    If Ganaron And .Score > 0 Then
                        ' Si está conectado
                        tUser = NameIndex(.username)
                        If IsValidUserRef(tUser) Then
                            ' Le damos el oro
                            Call WriteConsoleMsg(tUser.ArrayIndex, PremioStr, e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + OroGanado
                            Call WriteUpdateGold(tUser.ArrayIndex)
                        End If
                    End If
                    .username = vbNullString
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

Sub InvasionSpawnNPC(ByVal Index As Integer)
    With Invasiones(Index)
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
            Dim SpawnPos As t_WorldPos
            SpawnPos.Map = .TopLeft.Map
            SpawnPos.x = RandomNumber(.TopLeft.x, .BottomRight.x)
            SpawnPos.y = RandomNumber(.TopLeft.y, .BottomRight.y)
            ' Obtenemos la dirección y coordenada en la que se encuentra la muralla
            Dim Heading As e_Heading
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
            Call ChangeNPCChar(.NPCsVivos(i), NpcList(.NPCsVivos(i)).Char.body, NpcList(.NPCsVivos(i)).Char.head, Heading)
            ' Guardamos información sobre el spawn
            NpcList(.NPCsVivos(i)).flags.InvasionIndex = Index
            NpcList(.NPCsVivos(i)).flags.SpawnBox = SpawnBox
            NpcList(.NPCsVivos(i)).flags.IndexInInvasion = i
        End If
    End With
End Sub

Public Sub MuereNpcInvasion(ByVal Index As Integer, ByVal NpcIndex As Integer)
    With Invasiones(Index)
        .NPCsVivos(NpcIndex) = 0
        .CantNPCs = .CantNPCs - 1
    End With
End Sub

Private Function String2Heading(str As String) As e_Heading
    Select Case LCase$(str)
        Case "norte": String2Heading = e_Heading.NORTH
        Case "sur": String2Heading = e_Heading.SOUTH
        Case "este": String2Heading = e_Heading.EAST
        Case "oeste": String2Heading = e_Heading.WEST
    End Select
End Function

Public Sub EnviarInfoInvasion(ByVal Index As Integer)
    With Invasiones(Index)
        Dim PorcentajeVida As Byte, PorcentajeTiempo As Byte
        PorcentajeVida = (.VidaMuralla / .MaxVidaMuralla) * 100
        Dim elapsedMs As Double
        elapsedMs = TicksElapsed(.TiempoDeInicio, GetTickCountRaw())
        PorcentajeTiempo = CByte(elapsedMs / (.Duracion * 600))
        Dim i As Integer, Mapa As Integer, j As Integer
        For i = 1 To UBound(.SpawnBoxes)
            Mapa = .SpawnBoxes(i).TopLeft.Map
            For j = 1 To ModAreas.ConnGroups(Mapa).CountEntrys
                Call WriteInvasionInfo(ModAreas.ConnGroups(Mapa).UserEntrys(j), Index, PorcentajeVida, PorcentajeTiempo)
            Next
        Next
    End With
End Sub

Public Sub HacerDañoMuralla(ByVal Index As Integer, ByVal Daño As Long)
    With Invasiones(Index)
        .VidaMuralla = .VidaMuralla - Daño
        If .VidaMuralla <= 0 Then
            Call FinalizarInvasion(Index)
        End If
    End With
End Sub

Public Sub SumarScoreInvasion(ByVal Index As Integer, ByVal UserIndex As Integer, ByVal Score As Long)
    With Invasiones(Index)
        Dim i       As Integer
        Dim tmpUser As t_TopInvasion
        ' Buscamos si estamos en el top
        For i = 1 To UBound(.Top10Users)
            If LenB(.Top10Users(i).username) = 0 Then
                ' Llegamos a un lugar vacío, entonces no está en el top
                Exit For
            ElseIf .Top10Users(i).username = UserList(UserIndex).name Then
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
                .username = UserList(UserIndex).name
                .Score = Score
            End With
        End If
    End With
End Sub
