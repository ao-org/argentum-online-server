Attribute VB_Name = "ModAreas"
Option Explicit
 
'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
   
    AreaReciveX As Integer
    AreaReciveY As Integer
   
    MinX As Integer '-!!!
    MinY As Integer '-!!!
   
    AreaID As Long
End Type
 
Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type
 
Public Const USER_NUEVO As Byte = 255


Private Const AREA_DIM As Byte = 12
 
'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte
 
Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte
 
Private AreasRecive(10) As Integer
 
Public ConnGroups() As ConnGroup
 
Public Sub InitAreas()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
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

    ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
    For LoopC = 1 To NumMaps
        ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
        
        If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
        ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
    Next LoopC
End Sub
 
Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim LoopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
    
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        
        For LoopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
            
            ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
            If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
            If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
        Next LoopC
        
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub
 
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, ByVal appear As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la función clave del sistema de areas... Es llamada al mover un user
'**************************************************************
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, Y As Long
    Dim TempInt As Long, Map As Long

    With UserList(UserIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
            MinY = MinY + AREA_DIM * 3 '+ 27
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 4 * AREA_DIM - 1 ' + 35
            MinX = MinX + AREA_DIM * 3 '+ 27
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX - AREA_DIM * 2) '- 18)
            .AreasInfo.MinY = CInt(MinY)
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ AREA_DIM) - 1) * AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            
            MinX = ((.Pos.x \ AREA_DIM) - 1) * AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
       
        Map = UserList(UserIndex).Pos.Map
       
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
       
        'Actualizamos!!!
        For x = MinX To MaxX
            For Y = MinY To MaxY
               
                '<<< User >>>
                If MapData(Map, x, Y).UserIndex Then
                   
                    TempInt = MapData(Map, x, Y).UserIndex
                   
                    If UserIndex <> TempInt Then
                        Call MakeUserChar(False, UserIndex, TempInt, Map, x, Y, 0)
                        Call MakeUserChar(False, TempInt, UserIndex, .Pos.Map, .Pos.x, .Pos.Y, appear)
                       
                        'Si el user estaba invisible le avisamos al nuevo cliente de eso
                        If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                            Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                        End If
                        If UserList(UserIndex).flags.invisible Or UserList(UserIndex).flags.Oculto Then
                            Call WriteSetInvisible(TempInt, UserList(UserIndex).Char.CharIndex, True)
                        End If
                       
                        Call FlushBuffer(TempInt)
                   
                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(False, UserIndex, UserIndex, Map, x, Y, appear)
                    End If
                End If
               
                '<<< Npc >>>
                If MapData(Map, x, Y).NpcIndex Then
                    Call MakeNPCChar(False, UserIndex, MapData(Map, x, Y).NpcIndex, Map, x, Y)
                 End If
                 
                '<<< Item >>>
                If MapData(Map, x, Y).ObjInfo.ObjIndex Then
                    TempInt = MapData(Map, x, Y).ObjInfo.ObjIndex
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        Call WriteObjectCreate(UserIndex, TempInt, x, Y)
                       
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, x, Y, MapData(Map, x, Y).Blocked)
                            Call Bloquear(False, UserIndex, x - 1, Y, MapData(Map, x - 1, Y).Blocked)
                        End If
                    End If
                End If
                
                
                If MapData(Map, x, Y).Blocked = 2 Then
                    Call Bloquear(False, UserIndex, x, Y, True)
                 End If
                
                
                

                ' If MapData(Map, x, y).Particula > 0 Then
                   ' Call WriteParticleFloorCreate(UserIndex, MapData(Map, x, y).Particula, MapData(Map, x, y).TimeParticula, Map, x, y)
                 'End If
            
                'If MapData(Map, x, y).Luz.Rango > 0 Then
                    'Call WriteLightFloorCreate(UserIndex, MapData(Map, x, y).Luz.Color, MapData(Map, x, y).Luz.Rango, Map, x, y)
               ' End If
            Next Y
        Next x
       
        'Precalculados :P
        TempInt = .Pos.x \ AREA_DIM
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
        TempInt = .Pos.Y \ AREA_DIM
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.Y)
    End With
End Sub
Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' Se llama cuando se mueve un Npc
'**************************************************************
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, x As Long, Y As Long
    Dim TempInt As Long
    Dim appear As Byte
    appear = 0
    
    With Npclist(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
        
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - AREA_DIM
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
            MinY = MinY + AREA_DIM * 3 '+ 27
            MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 4 * AREA_DIM - 1 ' + 35
            MinX = MinX + AREA_DIM * 3 '+ 27
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            .AreasInfo.MinX = CInt(MinX - AREA_DIM * 2) '- 18)
            .AreasInfo.MinY = CInt(MinY)
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ AREA_DIM) - 1) * AREA_DIM
            MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            
            MinX = ((.Pos.x \ AREA_DIM) - 1) * AREA_DIM
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
        If MapInfo(.Pos.Map).NumUsers <> 0 Then
            For x = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(.Pos.Map, x, Y).UserIndex Then _
                        Call MakeNPCChar(False, MapData(.Pos.Map, x, Y).UserIndex, NpcIndex, .Pos.Map, .Pos.x, .Pos.Y)
                Next Y
            Next x
        End If
        
        'Precalculados :P
        TempInt = .Pos.x \ AREA_DIM
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
        TempInt = .Pos.Y \ AREA_DIM
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
        .AreasInfo.AreaID = AreasInfo(.Pos.x, .Pos.Y)
    End With
End Sub
 
Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    Dim LoopC As Long
   
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
   
    If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
    End If
End Sub
 
Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal appear As Byte = 0)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************
    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long

   
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
            ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long
        End If
       
        ConnGroups(Map).UserEntrys(TempVal) = UserIndex
    End If
   
    'Update user
    UserList(UserIndex).AreasInfo.AreaID = 0
   
    UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
    UserList(UserIndex).AreasInfo.AreaReciveX = 0
    UserList(UserIndex).AreasInfo.AreaReciveY = 0
   
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, appear)
End Sub
 
Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Npclist(NpcIndex).AreasInfo.AreaID = 0
   
    Npclist(NpcIndex).AreasInfo.AreaPerteneceX = 0
    Npclist(NpcIndex).AreasInfo.AreaPerteneceY = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveX = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveY = 0
   
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub

