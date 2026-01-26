Attribute VB_Name = "ModAreas"
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
'
Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type t_AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
    AreaReciveX As Integer
    AreaReciveY As Integer
    MinX As Integer '-!!!
    MinY As Integer '-!!!
    AreaID As Long
End Type

Public Type t_ConnGroup
    CountEntrys As Integer
    OptValue As Long
    UserEntrys() As Integer
    ' Ai uses user entry to search for nearby players in map
    'We cant to do the same with Npc that can be targeted by an AI becasue
    'looping about 10 char to se if they are in range is faster than lopping for 18x18 squares of vision
    RegisteredNpc As Integer
    NpcForAi() As Integer
End Type

Public Const USER_NUEVO               As Byte = 255
Private Const AREA_DIM                As Byte = 12
'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay                        As Byte
Private CurHour                       As Byte
Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100)           As Byte
Private AreasRecive(10)               As Integer
Public ConnGroups()                   As t_ConnGroup
 
Public Sub InitAreas()
    On Error GoTo InitAreas_Err
    Dim LoopC As Long
    Dim LoopX As Long
    ' Setup areas...
    For LoopC = 0 To 9
        AreasRecive(LoopC) = (2 ^ LoopC) Or IIf(LoopC <> 0, 2 ^ (LoopC - 1), 0) Or IIf(LoopC <> AREA_DIM, 2 ^ (LoopC + 1), 0)
    Next LoopC
    For LoopC = 1 To 100
        For LoopX = 1 To 100
            'Usamos 81 IDs de area para saber si pasasamos de area "más rápido"
            AreasInfo(LoopC, LoopX) = (LoopC \ AREA_DIM + 1) * (LoopX \ AREA_DIM + 1)
        Next LoopX
    Next LoopC
    'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
    ReDim ConnGroups(1 To NumMaps) As t_ConnGroup
    For LoopC = 1 To NumMaps
        ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
        If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
        ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Integer
        ReDim ConnGroups(LoopC).NpcForAi(0 To 1) As Integer
    Next LoopC
    Exit Sub
InitAreas_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.InitAreas", Erl)
End Sub
 
Public Sub AreasOptimizacion()
    On Error GoTo AreasOptimizacion_Err
    'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
    Dim LoopC            As Long
    Dim tCurDay          As Byte
    Dim tCurHour         As Byte
    Dim EntryValue       As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        For LoopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
            ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
            If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
            If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Integer
        Next LoopC
        CurDay = tCurDay
        CurHour = tCurHour
    End If
    Call PerformTimeLimitCheck(PerformanceTimer, "ModAreas.AreasOptimizacion")
    Exit Sub
AreasOptimizacion_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.AreasOptimizacion", Erl)
End Sub
 
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal head As Byte, ByVal appear As Byte, Optional ByVal Muerto As Byte = 0)
    On Error GoTo CheckUpdateNeededUser_Err
    'Es la función clave del sistema de areas... Es llamada al mover un user
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y) And Muerto = 0 Then Exit Sub
    Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, y As Long
    Dim TempInt As Long, Map As Long
    With UserList(UserIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        If head = e_Heading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = e_Heading.SOUTH Then
            MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
            MinY = MinY + AREA_DIM * 3 '+ 27
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        ElseIf head = e_Heading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = e_Heading.EAST Then
            MaxX = MinX + 4 * AREA_DIM - 1 ' + 35
            MinX = MinX + AREA_DIM * 3 '+ 27
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX - AREA_DIM * 2) '- 18)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.pos.y \ AREA_DIM) - 1) * AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            MinX = ((.pos.x \ AREA_DIM) - 1) * AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = 5 Then
            MaxX = MinX + AREA_DIM * 3 - 1
            MaxY = MinY + AREA_DIM * 3 - 1
        End If
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        Map = UserList(UserIndex).pos.Map
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
        'Actualizamos!!!
        For x = MinX To MaxX
            For y = MinY To MaxY
                '<<< User >>>
                If MapData(Map, x, y).UserIndex Then
                    TempInt = MapData(Map, x, y).UserIndex
                    If UserIndex <> TempInt Then
                        'NOTIFICO AL USUARIO QUE ESTABA EN EL AREA
                        Call NotifyUser(TempInt, UserIndex)
                        'NOTIFICO AL USUARIO QUE LLEGA AL AREA
                        Call NotifyUser(UserIndex, TempInt)
                    ElseIf head = USER_NUEVO Then
                        Call MakeUserChar(False, UserIndex, UserIndex, Map, x, y, appear)
                    End If
                End If
                '<<< Npc >>>
                If MapData(Map, x, y).NpcIndex Then
                    Call MakeNPCChar(False, UserIndex, MapData(Map, x, y).NpcIndex, Map, x, y)
                End If
                '<<< Item >>>
                If MapData(Map, x, y).ObjInfo.ObjIndex Then
                    TempInt = MapData(Map, x, y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        Call WriteObjectCreate(UserIndex, TempInt, MapData(Map, x, y).ObjInfo.amount, x, y)
                        If ObjData(TempInt).OBJType = e_OBJType.otDoors And InMapBounds(Map, x, y) Then
                            Call MostrarBloqueosPuerta(False, UserIndex, x, y)
                        End If
                    End If
                End If
                If Not MapData(Map, x, y).Trap Is Nothing Then
                    Call WriteUpdateTrapState(UserIndex, 1, x, y)
                End If
                ' Bloqueo GM
                If (MapData(Map, x, y).Blocked And e_Block.GM) <> 0 Then
                    Call Bloquear(False, UserIndex, x, y, e_Block.ALL_SIDES)
                End If
            Next y
        Next x
        'Precalculados :P
        TempInt = .pos.x \ AREA_DIM
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        TempInt = .pos.y \ AREA_DIM
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        .AreasInfo.AreaID = AreasInfo(.pos.x, .pos.y)
    End With
    Exit Sub
CheckUpdateNeededUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.CheckUpdateNeededUser", Erl)
End Sub

Private Sub NotifyUser(ByVal UserNotificado As Integer, ByVal UserIngresante As Integer)
    Dim sendChar As Boolean
    sendChar = True
    With UserList(UserNotificado)
        If UserList(UserIngresante).flags.AdminInvisible = 1 Then
            If Not EsGM(UserNotificado) Or CompararPrivilegios(.flags.Privilegios, UserList(UserIngresante).flags.Privilegios) <= 0 Then
                sendChar = False
            End If
        ElseIf UserList(UserNotificado).flags.Muerto = 1 And MapInfo(.pos.Map).Seguro = 0 And (UserList(UserNotificado).GuildIndex = 0 Or UserList(UserNotificado).GuildIndex <> _
                UserList(UserIngresante).GuildIndex Or modGuilds.NivelDeClan(UserList(UserIngresante).GuildIndex) < RequiredGuildLevelSeeInvisible) And Not IsSet(UserList(UserIngresante).flags.StatusMask, _
                e_StatusMask.eTalkToDead) Then
            sendChar = False
        End If
        If sendChar Then
            Call MakeUserChar(False, UserNotificado, UserIngresante, UserList(UserIngresante).pos.Map, UserList(UserIngresante).pos.x, UserList(UserIngresante).pos.y, 0)
            If UserList(UserIngresante).flags.invisible Or UserList(UserIngresante).flags.Oculto Then
                Call WriteSetInvisible(UserNotificado, UserIngresante, True)
            End If
        End If
    End With
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal head As Byte)
    On Error GoTo CheckUpdateNeededNpc_Err
    ' Se llama cuando se mueve un Npc
    If NpcList(NpcIndex).AreasInfo.AreaID = AreasInfo(NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y) Then Exit Sub
    Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, y As Long
    Dim TempInt As Long
    Dim appear  As Byte
    appear = 0
    With NpcList(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        If head = e_Heading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = e_Heading.SOUTH Then
            MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
            MinY = MinY + AREA_DIM * 3 '+ 27
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        ElseIf head = e_Heading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = e_Heading.EAST Then
            MaxX = MinX + 4 * AREA_DIM - 1 ' + 35
            MinX = MinX + AREA_DIM * 3 '+ 27
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX - AREA_DIM * 2) '- 18)
            .AreasInfo.MinY = CInt(MinY)
        ElseIf head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.pos.y \ AREA_DIM) - 1) * AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            MinX = ((.pos.x \ AREA_DIM) - 1) * AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
            appear = 0
        End If
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
        'Actualizamos!!!
        If MapInfo(.pos.Map).NumUsers <> 0 Then
            For x = MinX To MaxX
                For y = MinY To MaxY
                    If MapData(.pos.Map, x, y).UserIndex Then Call MakeNPCChar(False, MapData(.pos.Map, x, y).UserIndex, NpcIndex, .pos.Map, .pos.x, .pos.y)
                Next y
            Next x
        End If
        'Precalculados :P
        TempInt = .pos.x \ AREA_DIM
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        TempInt = .pos.y \ AREA_DIM
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        .AreasInfo.AreaID = AreasInfo(.pos.x, .pos.y)
    End With
    Exit Sub
CheckUpdateNeededNpc_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.CheckUpdateNeededNpc", Erl)
End Sub
 
Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)
    On Error GoTo QuitarUser_Err
    Dim TempVal As Long
    Dim LoopC   As Long
    'Search for the user
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(LoopC) = UserIndex Then Exit For
    Next LoopC
    'Char not found
    If LoopC > ConnGroups(Map).CountEntrys Then Exit Sub
    'Remove from old map
    ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
    TempVal = ConnGroups(Map).CountEntrys
    'Move list back
    For LoopC = LoopC To TempVal
        ConnGroups(Map).UserEntrys(LoopC) = ConnGroups(Map).UserEntrys(LoopC + 1)
    Next LoopC
    Exit Sub
QuitarUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.QuitarUser", Erl)
End Sub
 
Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal appear As Byte = 0)
    On Error GoTo AgregarUser_Err
    'Now the method checks for repetead users instead of trusting parameters.
    'If the character is new to the map, update it
    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i       As Long
    If Not MapaValido(Map) Then Exit Sub
    EsNuevo = True
    'Prevent adding repeated users
    For i = 1 To ConnGroups(Map).CountEntrys
        If ConnGroups(Map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
        TempVal = ConnGroups(Map).CountEntrys
        If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Integer
        End If
        ConnGroups(Map).UserEntrys(TempVal) = UserIndex
    End If
    'Update user
    UserList(UserIndex).AreasInfo.AreaID = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
    UserList(UserIndex).AreasInfo.AreaReciveX = 0
    UserList(UserIndex).AreasInfo.AreaReciveY = 0
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, appear, 1)
    Else
        Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, appear)
    End If
    Exit Sub
AgregarUser_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.AgregarUser", Erl)
End Sub
 
Public Sub AgregarNpc(ByVal NpcIndex As Integer)
    On Error GoTo AgregarNpc_Err
    With NpcList(NpcIndex)
        .AreasInfo.AreaID = 0
        .AreasInfo.AreaPerteneceX = 0
        .AreasInfo.AreaPerteneceY = 0
        .AreasInfo.AreaReciveX = 0
        .AreasInfo.AreaReciveY = 0
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eConsideredByMapAi) Then
            If ConnGroups(.pos.Map).RegisteredNpc = UBound(ConnGroups(.pos.Map).NpcForAi) Then
                ReDim Preserve ConnGroups(.pos.Map).NpcForAi(0 To ConnGroups(.pos.Map).RegisteredNpc + 1) As Integer
            End If
            ConnGroups(.pos.Map).NpcForAi(ConnGroups(.pos.Map).RegisteredNpc) = NpcIndex
            ConnGroups(.pos.Map).RegisteredNpc = ConnGroups(.pos.Map).RegisteredNpc + 1
        End If
    End With
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
    Exit Sub
AgregarNpc_Err:
    Call TraceError(Err.Number, Err.Description, "ModAreas.AgregarNpc", Erl)
End Sub

Public Sub RemoveNpc(ByVal NpcIndex As Integer)
    Dim i As Integer
    With NpcList(NpcIndex)
        If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eConsideredByMapAi) Then
            For i = 0 To ConnGroups(.pos.Map).RegisteredNpc
                If ConnGroups(.pos.Map).NpcForAi(i) = NpcIndex Then
                    ConnGroups(.pos.Map).NpcForAi(i) = ConnGroups(.pos.Map).NpcForAi(ConnGroups(.pos.Map).RegisteredNpc - 1)
                    ConnGroups(.pos.Map).NpcForAi(ConnGroups(.pos.Map).RegisteredNpc - 1) = 0
                    ConnGroups(.pos.Map).RegisteredNpc = ConnGroups(.pos.Map).RegisteredNpc - 1
                End If
            Next i
        End If
    End With
End Sub
