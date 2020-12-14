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
 
Public Const USER_NUEVO               As Byte = 255

Private Const AREA_DIM                As Byte = 12
 
'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay                        As Byte

Private CurHour                       As Byte
 
Private AreasInfo(1 To 100, 1 To 100) As Byte

Private PosToArea(1 To 100)           As Byte
 
Private AreasRecive(10)               As Integer
 
Public ConnGroups()                   As ConnGroup
 
Public Sub InitAreas()
        
        On Error GoTo InitAreas_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC As Long

        Dim LoopX As Long

        ' Setup areas...
100     For LoopC = 0 To 9
102         AreasRecive(LoopC) = (2 ^ LoopC) Or IIf(LoopC <> 0, 2 ^ (LoopC - 1), 0) Or IIf(LoopC <> AREA_DIM, 2 ^ (LoopC + 1), 0)
104     Next LoopC
    
106     For LoopC = 1 To 100
108         For LoopX = 1 To 100
                'Usamos 81 IDs de area para saber si pasasamos de area "más rápido"
110             AreasInfo(LoopC, LoopX) = (LoopC \ AREA_DIM + 1) * (LoopX \ AREA_DIM + 1)
112         Next LoopX
114     Next LoopC

        'Setup AutoOptimizacion de areas
116     CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
118     CurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece

120     ReDim ConnGroups(1 To NumMaps) As ConnGroup
    
122     For LoopC = 1 To NumMaps
124         ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
        
126         If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
128         ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
130     Next LoopC

132     Call generateMatrix(MATRIX_INITIAL_MAP)
        
        Exit Sub

InitAreas_Err:
134     Call RegistrarError(Err.Number, Err.description, "ModAreas.InitAreas", Erl)
136     Resume Next
        
End Sub
 
Public Sub AreasOptimizacion()
        
        On Error GoTo AreasOptimizacion_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Es la función de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
        '**************************************************************
        Dim LoopC      As Long

        Dim tCurDay    As Byte

        Dim tCurHour   As Byte

        Dim EntryValue As Long
    
100     If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(Time) \ 3)) Then
        
102         tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
104         tCurHour = Fix(Hour(Time) \ 3) 'A ke parte de la hora pertenece
        
106         For LoopC = 1 To NumMaps
108             EntryValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
110             Call WriteVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
            
112             ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, tCurDay & "-" & tCurHour))

114             If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
116             If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
118         Next LoopC
        
120         CurDay = tCurDay
122         CurHour = tCurHour

        End If

        
        Exit Sub

AreasOptimizacion_Err:
124     Call RegistrarError(Err.Number, Err.description, "ModAreas.AreasOptimizacion", Erl)
126     Resume Next
        
End Sub
 
Public Sub CheckUpdateNeededUser(ByVal Userindex As Integer, ByVal Head As Byte, ByVal appear As Byte)
        
        On Error GoTo CheckUpdateNeededUser_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Es la función clave del sistema de areas... Es llamada al mover un user
        '**************************************************************
100     If UserList(Userindex).AreasInfo.AreaID = AreasInfo(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) Then Exit Sub
    
        Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt As Long, Map As Long

102     With UserList(Userindex)
104         MinX = .AreasInfo.MinX
106         MinY = .AreasInfo.MinY
        
108         If Head = eHeading.NORTH Then
110             MaxY = MinY - 1
112             MinY = MinY - AREA_DIM
114             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
116             .AreasInfo.MinX = CInt(MinX)
118             .AreasInfo.MinY = CInt(MinY)
        
120         ElseIf Head = eHeading.SOUTH Then
122             MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
124             MinY = MinY + AREA_DIM * 3 '+ 27
126             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
128             .AreasInfo.MinX = CInt(MinX)
130             .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        
132         ElseIf Head = eHeading.WEST Then
134             MaxX = MinX - 1
136             MinX = MinX - AREA_DIM
138             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
140             .AreasInfo.MinX = CInt(MinX)
142             .AreasInfo.MinY = CInt(MinY)
        
144         ElseIf Head = eHeading.EAST Then
146             MaxX = MinX + 4 * AREA_DIM - 1 ' + 35
148             MinX = MinX + AREA_DIM * 3 '+ 27
150             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
152             .AreasInfo.MinX = CInt(MinX - AREA_DIM * 2) '- 18)
154             .AreasInfo.MinY = CInt(MinY)
           
156         ElseIf Head = USER_NUEVO Then
                'Esto pasa por cuando cambiamos de mapa o logeamos...
158             MinY = ((.Pos.Y \ AREA_DIM) - 1) * AREA_DIM
160             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            
162             MinX = ((.Pos.X \ AREA_DIM) - 1) * AREA_DIM
164             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            
166             .AreasInfo.MinX = CInt(MinX)
168             .AreasInfo.MinY = CInt(MinY)

            End If
        
170         If MinY < 1 Then MinY = 1
172         If MinX < 1 Then MinX = 1
174         If MaxY > 100 Then MaxY = 100
176         If MaxX > 100 Then MaxX = 100
       
178         Map = UserList(Userindex).Pos.Map
       
            'Esto es para ke el cliente elimine lo "fuera de area..."
180         Call WriteAreaChanged(Userindex)
       
            'Actualizamos!!!
182         For X = MinX To MaxX
184             For Y = MinY To MaxY
               
                    '<<< User >>>
186                 If MapData(Map, X, Y).Userindex Then
                   
188                     TempInt = MapData(Map, X, Y).Userindex
                   
190                     If Userindex <> TempInt Then
                        
192                         If UserList(Userindex).flags.AdminInvisible = 0 Or EsGM(TempInt) Then
194                             Call MakeUserChar(False, TempInt, Userindex, .Pos.Map, .Pos.X, .Pos.Y, 0)
                            
196                             If UserList(Userindex).flags.invisible Or UserList(Userindex).flags.Oculto Then
198                                 Call WriteSetInvisible(TempInt, UserList(Userindex).Char.CharIndex, True)
                                End If
                            
                            End If
                            
200                         If UserList(TempInt).flags.AdminInvisible = 0 Or EsGM(Userindex) Then
202                             Call MakeUserChar(False, Userindex, TempInt, Map, X, Y, appear)
                            
                                'Si el user estaba invisible le avisamos al nuevo cliente de eso
204                             If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
206                                 Call WriteSetInvisible(Userindex, UserList(TempInt).Char.CharIndex, True)
                                End If

                            End If

208                     ElseIf Head = USER_NUEVO Then
210                         Call MakeUserChar(False, Userindex, Userindex, Map, X, Y, appear)

                        End If

                    End If
               
                    '<<< Npc >>>
212                 If MapData(Map, X, Y).NpcIndex Then
214                     Call MakeNPCChar(False, Userindex, MapData(Map, X, Y).NpcIndex, Map, X, Y)

                    End If
                 
                    '<<< Item >>>
216                 If MapData(Map, X, Y).ObjInfo.ObjIndex Then
218                     TempInt = MapData(Map, X, Y).ObjInfo.ObjIndex

220                     If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
222                         Call WriteObjectCreate(Userindex, TempInt, X, Y)
                       
224                         If ObjData(TempInt).OBJType = eOBJType.otPuertas And InMapBounds(Map, X, Y) Then
226                             Call MostrarBloqueosPuerta(False, Userindex, X, Y)
                            End If

                        End If

                    End If

                    ' Bloqueo GM
228                 If (MapData(Map, X, Y).Blocked And eBlock.GM) <> 0 Then
230                     Call Bloquear(False, Userindex, X, Y, eBlock.ALL_SIDES)
                    End If

                    ' If MapData(Map, x, y).Particula > 0 Then
                    ' Call WriteParticleFloorCreate(UserIndex, MapData(Map, x, y).Particula, MapData(Map, x, y).TimeParticula, Map, x, y)
                    'End If
            
                    'If MapData(Map, x, y).Luz.Rango > 0 Then
                    'Call WriteLightFloorCreate(UserIndex, MapData(Map, x, y).Luz.Color, MapData(Map, x, y).Luz.Rango, Map, x, y)
                    ' End If
232             Next Y
234         Next X
       
            'Precalculados :P
236         TempInt = .Pos.X \ AREA_DIM
238         .AreasInfo.AreaReciveX = AreasRecive(TempInt)
240         .AreasInfo.AreaPerteneceX = 2 ^ TempInt
        
242         TempInt = .Pos.Y \ AREA_DIM
244         .AreasInfo.AreaReciveY = AreasRecive(TempInt)
246         .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
248         .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)

        End With

        
        Exit Sub

CheckUpdateNeededUser_Err:
250     Call RegistrarError(Err.Number, Err.description, "ModAreas.CheckUpdateNeededUser", Erl)
252     Resume Next
        
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
        
        On Error GoTo CheckUpdateNeededNpc_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        ' Se llama cuando se mueve un Npc
        '**************************************************************
100     If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
    
        Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt As Long

        Dim appear  As Byte

102     appear = 0
    
104     With Npclist(NpcIndex)
106         MinX = .AreasInfo.MinX
108         MinY = .AreasInfo.MinY
        
110         If Head = eHeading.NORTH Then
112             MaxY = MinY - 1
114             MinY = MinY - AREA_DIM
116             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
118             .AreasInfo.MinX = CInt(MinX)
120             .AreasInfo.MinY = CInt(MinY)
        
122         ElseIf Head = eHeading.SOUTH Then
124             MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
126             MinY = MinY + AREA_DIM * 3 '+ 27
128             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
130             .AreasInfo.MinX = CInt(MinX)
132             .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        
134         ElseIf Head = eHeading.WEST Then
136             MaxX = MinX - 1
138             MinX = MinX - AREA_DIM
140             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
142             .AreasInfo.MinX = CInt(MinX)
144             .AreasInfo.MinY = CInt(MinY)
        
146         ElseIf Head = eHeading.EAST Then
148             MaxX = MinX + 4 * AREA_DIM - 1 ' + 35
150             MinX = MinX + AREA_DIM * 3 '+ 27
152             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
154             .AreasInfo.MinX = CInt(MinX - AREA_DIM * 2) '- 18)
156             .AreasInfo.MinY = CInt(MinY)
           
158         ElseIf Head = USER_NUEVO Then
                'Esto pasa por cuando cambiamos de mapa o logeamos...
160             MinY = ((.Pos.Y \ AREA_DIM) - 1) * AREA_DIM
162             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
            
164             MinX = ((.Pos.X \ AREA_DIM) - 1) * AREA_DIM
166             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
            
168             .AreasInfo.MinX = CInt(MinX)
170             .AreasInfo.MinY = CInt(MinY)
            
172             appear = 0

            End If
        
174         If MinY < 1 Then MinY = 1
176         If MinX < 1 Then MinX = 1
178         If MaxY > 100 Then MaxY = 100
180         If MaxX > 100 Then MaxX = 100
        
            'Actualizamos!!!
182         If MapInfo(.Pos.Map).NumUsers <> 0 Then

184             For X = MinX To MaxX
186                 For Y = MinY To MaxY

188                     If MapData(.Pos.Map, X, Y).Userindex Then Call MakeNPCChar(False, MapData(.Pos.Map, X, Y).Userindex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)
190                 Next Y
192             Next X

            End If
        
            'Precalculados :P
194         TempInt = .Pos.X \ AREA_DIM
196         .AreasInfo.AreaReciveX = AreasRecive(TempInt)
198         .AreasInfo.AreaPerteneceX = 2 ^ TempInt
            
200         TempInt = .Pos.Y \ AREA_DIM
202         .AreasInfo.AreaReciveY = AreasRecive(TempInt)
204         .AreasInfo.AreaPerteneceY = 2 ^ TempInt
        
206         .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)

        End With

        
        Exit Sub

CheckUpdateNeededNpc_Err:
208     Call RegistrarError(Err.Number, Err.description, "ModAreas.CheckUpdateNeededNpc", Erl)
210     Resume Next
        
End Sub
 
Public Sub QuitarUser(ByVal Userindex As Integer, ByVal Map As Integer)
        
        On Error GoTo QuitarUser_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim TempVal As Long

        Dim LoopC   As Long
   
        'Search for the user
100     For LoopC = 1 To ConnGroups(Map).CountEntrys

102         If ConnGroups(Map).UserEntrys(LoopC) = Userindex Then Exit For
104     Next LoopC
   
        'Char not found
106     If LoopC > ConnGroups(Map).CountEntrys Then Exit Sub
   
        'Remove from old map
108     ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys - 1
110     TempVal = ConnGroups(Map).CountEntrys
   
        'Move list back
112     For LoopC = LoopC To TempVal
114         ConnGroups(Map).UserEntrys(LoopC) = ConnGroups(Map).UserEntrys(LoopC + 1)
116     Next LoopC
   
118     If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim?
120         ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long

        End If

        
        Exit Sub

QuitarUser_Err:
122     Call RegistrarError(Err.Number, Err.description, "ModAreas.QuitarUser", Erl)
124     Resume Next
        
End Sub
 
Public Sub AgregarUser(ByVal Userindex As Integer, ByVal Map As Integer, Optional ByVal appear As Byte = 0)
        
        On Error GoTo AgregarUser_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: 04/01/2007
        'Modified by Juan Martín Sotuyo Dodero (Maraxus)
        '   - Now the method checks for repetead users instead of trusting parameters.
        '   - If the character is new to the map, update it
        '**************************************************************
        Dim TempVal As Long

        Dim EsNuevo As Boolean

        Dim i       As Long
   
100     If Not MapaValido(Map) Then Exit Sub
   
102     EsNuevo = True
   
        'Prevent adding repeated users
104     For i = 1 To ConnGroups(Map).CountEntrys

106         If ConnGroups(Map).UserEntrys(i) = Userindex Then
108             EsNuevo = False
                Exit For

            End If

110     Next i
   
112     If EsNuevo Then
            'Update map and connection groups data
114         ConnGroups(Map).CountEntrys = ConnGroups(Map).CountEntrys + 1
116         TempVal = ConnGroups(Map).CountEntrys
       
118         If TempVal > ConnGroups(Map).OptValue Then 'Nescesito Redim
120             ReDim Preserve ConnGroups(Map).UserEntrys(1 To TempVal) As Long

            End If
       
122         ConnGroups(Map).UserEntrys(TempVal) = Userindex

        End If
   
        'Update user
124     UserList(Userindex).AreasInfo.AreaID = 0
   
126     UserList(Userindex).AreasInfo.AreaPerteneceX = 0
128     UserList(Userindex).AreasInfo.AreaPerteneceY = 0
130     UserList(Userindex).AreasInfo.AreaReciveX = 0
132     UserList(Userindex).AreasInfo.AreaReciveY = 0
   
134     Call CheckUpdateNeededUser(Userindex, USER_NUEVO, appear)

        
        Exit Sub

AgregarUser_Err:
136     Call RegistrarError(Err.Number, Err.description, "ModAreas.AgregarUser", Erl)
138     Resume Next
        
End Sub
 
Public Sub AgregarNpc(ByVal NpcIndex As Integer)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        
        On Error GoTo AgregarNpc_Err
        
100     Npclist(NpcIndex).AreasInfo.AreaID = 0
   
102     Npclist(NpcIndex).AreasInfo.AreaPerteneceX = 0
104     Npclist(NpcIndex).AreasInfo.AreaPerteneceY = 0
106     Npclist(NpcIndex).AreasInfo.AreaReciveX = 0
108     Npclist(NpcIndex).AreasInfo.AreaReciveY = 0
   
110     Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)

        
        Exit Sub

AgregarNpc_Err:
112     Call RegistrarError(Err.Number, Err.description, "ModAreas.AgregarNpc", Erl)
114     Resume Next
        
End Sub

