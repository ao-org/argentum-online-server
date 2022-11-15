Attribute VB_Name = "ModInvasion"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
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
    UserName As String
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
100 Set File = New clsIniManager

102     Call File.Initialize(DatPath & "Eventos.dat")
    
        Dim CantInvasiones As Integer
104     CantInvasiones = val(File.GetValue("Invasiones", "Cantidad"))

106     If CantInvasiones <= 0 Then
108         ReDim Invasiones(0)
            frmMain.Invasion.Enabled = False
            Exit Sub
        End If

110     ReDim Invasiones(1 To CantInvasiones)
    
        Dim i As Integer, j As Integer, nombre As String, tmpStr As String, Fields() As String
    
112     For i = 1 To CantInvasiones
        
114         nombre = File.GetValue("Invasiones", "Invasion" & i)
        
116         With Invasiones(i)
    
118             .MaxVidaMuralla = val(File.GetValue(nombre, "MaxVidaMuralla"))
120             .MaxNPCs = val(File.GetValue(nombre, "MaxNPCs"))
122             .aviso = File.GetValue(nombre, "Aviso")
124             .AvisarTiempo = val(File.GetValue(nombre, "AvisarTiempo"))
126             .RepetirAviso = val(File.GetValue(nombre, "RepetirAviso"))
128             .Desc = File.GetValue(nombre, "Desc")
130             .RepetirDesc = val(File.GetValue(nombre, "RepetirDesc"))
132             .IntervaloSpawn = val(File.GetValue(nombre, "IntervaloSpawn"))
134             .Duracion = val(File.GetValue(nombre, "Duracion"))
136             .Intervalo = val(File.GetValue(nombre, "Intervalo"))
138             .TimerInvasion = val(File.GetValue(nombre, "Offset"))
140             .MensajeGanaron = File.GetValue(nombre, "MensajeGanaron")
142             .MensajePerdieron = File.GetValue(nombre, "MensajePerdieron")
            
144             If .MaxNPCs <= 0 Then Exit Sub

146             ReDim .NPCsVivos(1 To .MaxNPCs)
            
148             tmpStr = File.GetValue(nombre, "NPCs")
150             If LenB(tmpStr) > 0 Then
152                 Fields = Split(tmpStr, "-")
                
154                 ReDim .NumNPCsSpawn(1 To UBound(Fields) + 1)
                
156                 For j = 1 To UBound(.NumNPCsSpawn)
158                     .NumNPCsSpawn(j) = val(Fields(j - 1))
                    Next
                Else
160                 ReDim .NumNPCsSpawn(0)
                End If
            
                Dim SpawnBoxes As Integer
162             SpawnBoxes = val(File.GetValue(nombre, "SpawnBoxes"))
            
164             If SpawnBoxes <= 0 Then Exit Sub
            
166             ReDim .SpawnBoxes(1 To SpawnBoxes)
            
168             For j = 1 To SpawnBoxes
170                 tmpStr = File.GetValue(nombre, "SpawnBox" & j)
                
172                 If LenB(tmpStr) > 0 Then
174                     Fields = Split(tmpStr, "-", 7)
                
176                     If UBound(Fields) = 6 Then
178                         With .SpawnBoxes(j)
                                ' Mapa
180                             .TopLeft.Map = val(Fields(0))
182                             .BottomRight.Map = .TopLeft.Map
                                ' TopLeft
184                             .TopLeft.X = val(Fields(1))
186                             .TopLeft.Y = val(Fields(2))
                                ' BottomRight
188                             .BottomRight.X = val(Fields(3))
190                             .BottomRight.Y = val(Fields(4))
                                ' Dirección de ataque
192                             .Heading = String2Heading(Fields(5))
194                             .CoordMuralla = val(Fields(6))
                                ' Calculamos las posiciones válidas de los NPCs
196                             .LegalBox.X1 = .TopLeft.X
198                             .LegalBox.Y1 = .TopLeft.Y
200                             .LegalBox.X2 = .BottomRight.X
202                             .LegalBox.Y2 = .BottomRight.Y

204                             Select Case .Heading
                                    Case e_Heading.NORTH: .LegalBox.Y1 = .CoordMuralla
206                                 Case e_Heading.SOUTH: .LegalBox.Y2 = .CoordMuralla
208                                 Case e_Heading.EAST: .LegalBox.X2 = .CoordMuralla
210                                 Case e_Heading.WEST: .LegalBox.X1 = .CoordMuralla
                                End Select
                            End With
                        End If
                    End If
                
                Next
            
            End With
        Next
        
        frmMain.Invasion.Enabled = True
        
212     Set File = Nothing
End Sub

Sub IniciarInvasion(ByVal Index As Integer)
        
        If UBound(Invasiones) = 0 Then Exit Sub
        
100     With Invasiones(Index)
    
102         .Activa = True
        
104         .VidaMuralla = .MaxVidaMuralla
        
106         .TiempoDeInicio = GetTickCount

            ' Enviamos info sobre la invasión a los usuarios en estos mapas
108         Call EnviarInfoInvasion(Index)
        
110         Call MensajeGlobal(.Desc, e_FontTypeNames.FONTTYPE_New_Eventos)
112         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(150, NO_3D_SOUND, NO_3D_SOUND))
    
        End With
    
End Sub

Sub FinalizarInvasion(ByVal Index As Integer)

100     With Invasiones(Index)
    
            Dim Ganaron As Boolean
        
102         If .VidaMuralla > 0 Then
104             Call MensajeGlobal(.MensajeGanaron, e_FontTypeNames.FONTTYPE_New_Eventos)
106             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
108             Ganaron = True
            Else
110             Call MensajeGlobal(.MensajePerdieron, e_FontTypeNames.FONTTYPE_New_Eventos)
112             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            End If

            ' Limpiamos flags
114         .Activa = False
116         .TimerRepetirAviso = 0
118         .TimerInvasion = 0
120         .TimerMostrarInfo = 0
122         .TimerRepetirDesc = 0
124         .TimerSpawn = 0

            ' Matamos los NPCs que quedaron
            Dim i As Integer
126         For i = 1 To UBound(.NPCsVivos)
128             If .NPCsVivos(i) Then
130                 Call QuitarNPC(.NPCsVivos(i), eClearInvasion)
132                 .NPCsVivos(i) = 0
                
134                 .CantNPCs = .CantNPCs - 1
136                 If .CantNPCs <= 0 Then Exit For
                End If
            Next
        
            ' Entregamos premios y limpiamos el top
            Dim tUser As t_UserReference, OroGanado As Long, PremioStr As String
    
138         OroGanado = 50000 * OroMult
140         PremioStr = "¡La ciudad te entrega " & PonerPuntos(OroGanado) & " monedas de oro por tu ayuda durante la invasión!"
        
142         For i = 1 To UBound(.Top10Users)
144             With .Top10Users(i)
        
146                 If LenB(.UserName) Then
148                     If Ganaron And .Score > 0 Then
                            ' Si está conectado
150                         tUser = NameIndex(.username)
152                         If IsValidUserRef(tUser) Then
                                ' Le damos el oro
154                             Call WriteConsoleMsg(tUser.ArrayIndex, PremioStr, e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
156                             UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + OroGanado
158                             Call WriteUpdateGold(tUser.ArrayIndex)
                            End If
                        End If
                    
160                     .UserName = vbNullString
162                     .Score = 0
                    End If
                End With
            Next

            ' Sacamos el cartel de la pantalla de todos
            Dim Mapa As Integer, j As Integer
164         For i = 1 To UBound(.SpawnBoxes)
166             Mapa = .SpawnBoxes(i).TopLeft.Map
            
168             For j = 1 To ModAreas.ConnGroups(Mapa).CountEntrys
170                 Call WriteInvasionInfo(ModAreas.ConnGroups(Mapa).UserEntrys(j), 0, 0, 0)
                Next
            Next

        End With

End Sub

Sub InvasionSpawnNPC(ByVal index As Integer)

100     With Invasiones(index)
    
            ' Si ya hay el máximo de NPCs, no spawneamos nada
102         If .CantNPCs >= .MaxNPCs Then Exit Sub
        
            ' Elegimos al azar el NPC a spawnear
            Dim NpcNumber As Integer
104         NpcNumber = .NumNPCsSpawn(RandomNumber(1, UBound(.NumNPCsSpawn)))
        
            ' Elegimos un área al azar (TODO: elegir con más probabilidad según más área cubra)
            Dim SpawnBox As Integer
106         SpawnBox = RandomNumber(1, UBound(.SpawnBoxes))
        
108         With .SpawnBoxes(SpawnBox)
        
                ' Elegimos un tile al azar dentro del área
                Dim SpawnPos As t_WorldPos
110             SpawnPos.Map = .TopLeft.Map
112             SpawnPos.X = RandomNumber(.TopLeft.X, .BottomRight.X)
114             SpawnPos.Y = RandomNumber(.TopLeft.Y, .BottomRight.Y)
            
                ' Obtenemos la dirección y coordenada en la que se encuentra la muralla
                Dim Heading As e_Heading
116             Heading = .Heading
        
            End With

            ' Buscamos un índice vacío en el array de NPCs
            Dim i As Integer
118         For i = 1 To UBound(.NPCsVivos)
120             If .NPCsVivos(i) = 0 Then Exit For
            Next

            ' Spawneamos el NPC
122         .NPCsVivos(i) = SpawnNpc(NpcNumber, SpawnPos, True, False)
        
124         Debug.Assert .NPCsVivos(i) <> 0

            ' Si pudimos spawnearlo
126         If .NPCsVivos(i) Then
128             .CantNPCs = .CantNPCs + 1
    
                ' Lo colocamos mirando en dirección a la muralla
130             Call ChangeNPCChar(.NPCsVivos(i), NpcList(.NPCsVivos(i)).Char.Body, NpcList(.NPCsVivos(i)).Char.Head, Heading)
            
                ' Guardamos información sobre el spawn
132             NpcList(.NPCsVivos(i)).flags.InvasionIndex = index
134             NpcList(.NPCsVivos(i)).flags.SpawnBox = SpawnBox
136             NpcList(.NPCsVivos(i)).flags.IndexInInvasion = i
            End If

        End With

End Sub

Public Sub MuereNpcInvasion(ByVal index As Integer, ByVal NpcIndex As Integer)

100     With Invasiones(index)
    
102         .NPCsVivos(NpcIndex) = 0
    
104         .CantNPCs = .CantNPCs - 1

        End With

End Sub

Private Function String2Heading(str As String) As e_Heading

100     Select Case LCase$(str)
            Case "norte": String2Heading = e_Heading.NORTH
102         Case "sur": String2Heading = e_Heading.SOUTH
104         Case "este": String2Heading = e_Heading.EAST
106         Case "oeste": String2Heading = e_Heading.WEST
        End Select

End Function

Public Sub EnviarInfoInvasion(ByVal index As Integer)

100     With Invasiones(index)
    
            Dim PorcentajeVida As Byte, PorcentajeTiempo As Byte
        
102         PorcentajeVida = (.VidaMuralla / .MaxVidaMuralla) * 100
        
104         PorcentajeTiempo = (GetTickCount - .TiempoDeInicio) / (.Duracion * 600)
    
            Dim i As Integer, Mapa As Integer, j As Integer
106         For i = 1 To UBound(.SpawnBoxes)
108             Mapa = .SpawnBoxes(i).TopLeft.Map
            
110             For j = 1 To ModAreas.ConnGroups(Mapa).CountEntrys
112                 Call WriteInvasionInfo(ModAreas.ConnGroups(Mapa).UserEntrys(j), index, PorcentajeVida, PorcentajeTiempo)
                Next
            Next
        
        End With

End Sub

Public Sub HacerDañoMuralla(ByVal Index As Integer, ByVal Daño As Long)
    
100     With Invasiones(index)
    
102         .VidaMuralla = .VidaMuralla - Daño
        
104         If .VidaMuralla <= 0 Then
106             Call FinalizarInvasion(index)
            End If
    
        End With
    
End Sub

Public Sub SumarScoreInvasion(ByVal index As Integer, ByVal UserIndex As Integer, ByVal Score As Long)
    
100     With Invasiones(index)
    
            Dim i As Integer
            Dim tmpUser As t_TopInvasion
        
            ' Buscamos si estamos en el top
102         For i = 1 To UBound(.Top10Users)
104             If LenB(.Top10Users(i).UserName) = 0 Then
                    ' Llegamos a un lugar vacío, entonces no está en el top
                    Exit For
            
106             ElseIf .Top10Users(i).UserName = UserList(UserIndex).name Then
                    ' Está en el top, así que le sumamos el puntaje
108                 .Top10Users(i).Score = .Top10Users(i).Score + Score
                
                    ' Revisamos si subió en el top
                    Dim j As Integer
110                 For j = i - 1 To 1 Step -1
                        ' Si el que está arriba tiene un puntaje menor, los cambiamos
112                     If .Top10Users(j).Score < .Top10Users(j + 1).Score Then
114                         tmpUser = .Top10Users(j)
116                         .Top10Users(j) = .Top10Users(j + 1)
118                         .Top10Users(j + 1) = tmpUser
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
120         For i = UBound(.Top10Users) To 1 Step -1
                ' Buscamos el lugar indicado
122             If .Top10Users(i).Score > Score Then
                    Exit For
                End If
            Next
        
            ' Si entró en el top
124         If i < UBound(.Top10Users) Then
                ' Movemos a los que le siguen
126             For j = UBound(.Top10Users) To i + 2
128                 .Top10Users(j) = .Top10Users(j - 1)
                Next
            
                ' Lo colocamos en la posición que le corresponde
130             With .Top10Users(i + 1)
132                 .UserName = UserList(UserIndex).name
134                 .Score = Score
                End With
            End If
    
        End With
    
End Sub
