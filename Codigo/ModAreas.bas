Attribute VB_Name = "ModAreas"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
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
 
Public ConnGroups()                   As t_ConnGroup
 
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

120     ReDim ConnGroups(1 To NumMaps) As t_ConnGroup
    
122     For LoopC = 1 To NumMaps
124         ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.ini", "Mapa" & LoopC, CurDay & "-" & CurHour))
        
126         If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
128         ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
130     Next LoopC
        
        Exit Sub

InitAreas_Err:
132     Call TraceError(Err.Number, Err.Description, "ModAreas.InitAreas", Erl)

        
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
124     Call TraceError(Err.Number, Err.Description, "ModAreas.AreasOptimizacion", Erl)

        
End Sub
 
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, ByVal appear As Byte)
        
        On Error GoTo CheckUpdateNeededUser_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Es la función clave del sistema de areas... Es llamada al mover un user
        '**************************************************************
100     If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then Exit Sub
    
        Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt As Long, Map As Long

102     With UserList(UserIndex)
104         MinX = .AreasInfo.MinX
106         MinY = .AreasInfo.MinY
        
108         If Head = e_Heading.NORTH Then
110             MaxY = MinY - 1
112             MinY = MinY - AREA_DIM
114             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
116             .AreasInfo.MinX = CInt(MinX)
118             .AreasInfo.MinY = CInt(MinY)
        
120         ElseIf Head = e_Heading.SOUTH Then
122             MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
124             MinY = MinY + AREA_DIM * 3 '+ 27
126             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
128             .AreasInfo.MinX = CInt(MinX)
130             .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        
132         ElseIf Head = e_Heading.WEST Then
134             MaxX = MinX - 1
136             MinX = MinX - AREA_DIM
138             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
140             .AreasInfo.MinX = CInt(MinX)
142             .AreasInfo.MinY = CInt(MinY)
        
144         ElseIf Head = e_Heading.EAST Then
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
       
178         Map = UserList(UserIndex).Pos.Map
       
            'Esto es para ke el cliente elimine lo "fuera de area..."
180         Call WriteAreaChanged(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If UserList(UserIndex).flags.GMMeSigue > 0 Then
                Call WriteAreaChanged(UserList(UserIndex).flags.GMMeSigue, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                Call WriteSendFollowingCharindex(UserList(UserIndex).flags.GMMeSigue, UserList(UserIndex).Char.charindex)
            End If
       
            'Actualizamos!!!
182         For X = MinX To MaxX
184             For Y = MinY To MaxY
               
                    '<<< User >>>
186                 If MapData(Map, X, Y).UserIndex Then
                   
188                     TempInt = MapData(Map, X, Y).UserIndex
190                     If UserIndex <> TempInt Then
                        
192                         If (UserList(UserIndex).flags.AdminInvisible = 0 Or EsGM(TempInt)) And (UserList(TempInt).flags.Muerto = 0 Or UserList(TempInt).GuildIndex = UserList(UserIndex).GuildIndex) Or (UserList(UserIndex).flags.Muerto = 1 And UserList(TempInt).flags.Muerto = 1) Then
                                If UserList(TempInt).flags.SigueUsuario = 0 Then
194                                 Call MakeUserChar(False, TempInt, UserIndex, .Pos.map, .Pos.X, .Pos.Y, 0)
                                End If
                                'If tmpGM > 0 Then Call MakeUserChar(False, tmpGM, UserIndex, .Pos.map, .Pos.X, .Pos.Y, 0)

196                             If UserList(UserIndex).flags.invisible Or UserList(UserIndex).flags.Oculto Then
198                                 Call WriteSetInvisible(TempInt, UserList(UserIndex).Char.CharIndex, True)
                                End If
                            
                            End If
                            
200                         If (UserList(TempInt).flags.AdminInvisible = 0 Or EsGM(UserIndex)) And (UserList(UserIndex).flags.Muerto = 0 Or UserList(TempInt).GuildIndex = UserList(UserIndex).GuildIndex) Or (UserList(UserIndex).flags.Muerto = 1 And UserList(TempInt).flags.Muerto = 1) Then
202                                 Call MakeUserChar(False, UserIndex, TempInt, map, X, Y, appear)

                                
                                'Si el user estaba invisible le avisamos al nuevo cliente de eso
204                             If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
206                                 Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                                End If

                            End If

208                     ElseIf Head = USER_NUEVO Then
210                         Call MakeUserChar(False, UserIndex, UserIndex, Map, X, Y, appear)
                        End If

                    End If
               
                    '<<< Npc >>>
212                 If MapData(Map, X, Y).NpcIndex Then
214                     Call MakeNPCChar(False, UserIndex, MapData(Map, X, Y).NpcIndex, Map, X, Y)
                       ' If tmpGM > 0 Then
                       '     Call MakeNPCChar(False, tmpGM, MapData(map, X, Y).NpcIndex, map, X, Y)
                       ' End If
                    End If
                 
                    '<<< Item >>>
216                 If MapData(Map, X, Y).ObjInfo.ObjIndex Then
218                     TempInt = MapData(Map, X, Y).ObjInfo.ObjIndex

220                     If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
222                         Call WriteObjectCreate(UserIndex, TempInt, MapData(Map, X, Y).ObjInfo.amount, X, Y)
                           ' If tmpGM > 0 Then Call WriteObjectCreate(tmpGM, TempInt, MapData(map, X, Y).ObjInfo.amount, X, Y)
                            
224                         If ObjData(TempInt).OBJType = e_OBJType.otPuertas And InMapBounds(Map, X, Y) Then
226                             Call MostrarBloqueosPuerta(False, UserIndex, X, Y)
                                'If tmpGM > 0 Then Call MostrarBloqueosPuerta(False, tmpGM, X, Y)
                            End If
                            
                        End If

                    End If

                    ' Bloqueo GM
228                 If (MapData(Map, X, Y).Blocked And e_Block.GM) <> 0 Then
230                     Call Bloquear(False, UserIndex, X, Y, e_Block.ALL_SIDES)
                    End If
                    
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
            
            'Es un gm que está siguiendo a un usuario
            If .flags.SigueUsuario > 0 Then
              .AreasInfo.AreaReciveX = UserList(.flags.SigueUsuario).AreasInfo.AreaReciveX
              .AreasInfo.AreaPerteneceX = UserList(.flags.SigueUsuario).AreasInfo.AreaPerteneceX
            
              .AreasInfo.AreaReciveY = UserList(.flags.SigueUsuario).AreasInfo.AreaReciveY
              .AreasInfo.AreaPerteneceY = UserList(.flags.SigueUsuario).AreasInfo.AreaPerteneceY
             
             .AreasInfo.AreaID = UserList(.flags.SigueUsuario).AreasInfo.AreaID
            
            End If
            
            'Es un usuario que está siendo seguido
            If .flags.GMMeSigue > 0 Then
                UserList(.flags.GMMeSigue).AreasInfo.AreaReciveX = .AreasInfo.AreaReciveX
                UserList(.flags.GMMeSigue).AreasInfo.AreaPerteneceX = .AreasInfo.AreaPerteneceX
                
                UserList(.flags.GMMeSigue).AreasInfo.AreaReciveY = .AreasInfo.AreaReciveY
                UserList(.flags.GMMeSigue).AreasInfo.AreaPerteneceY = .AreasInfo.AreaPerteneceY
             
                UserList(.flags.GMMeSigue).AreasInfo.AreaID = .AreasInfo.AreaID
            
            End If
        End With

        
        Exit Sub

CheckUpdateNeededUser_Err:
250     Call TraceError(Err.Number, Err.Description, "ModAreas.CheckUpdateNeededUser", Erl)

        
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
        
        On Error GoTo CheckUpdateNeededNpc_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        ' Se llama cuando se mueve un Npc
        '**************************************************************
100     If NpcList(NpcIndex).AreasInfo.AreaID = AreasInfo(NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y) Then Exit Sub
    
        Dim MinX    As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long

        Dim TempInt As Long

        Dim appear  As Byte

102     appear = 0
    
104     With NpcList(NpcIndex)
106         MinX = .AreasInfo.MinX
108         MinY = .AreasInfo.MinY
        
110         If Head = e_Heading.NORTH Then
112             MaxY = MinY - 1
114             MinY = MinY - AREA_DIM
116             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
118             .AreasInfo.MinX = CInt(MinX)
120             .AreasInfo.MinY = CInt(MinY)
        
122         ElseIf Head = e_Heading.SOUTH Then
124             MaxY = MinY + 4 * AREA_DIM - 1 ' + 35
126             MinY = MinY + AREA_DIM * 3 '+ 27
128             MaxX = MinX + AREA_DIM * 3 - 1 '+ 26
130             .AreasInfo.MinX = CInt(MinX)
132             .AreasInfo.MinY = CInt(MinY - AREA_DIM * 2) '- 18)
        
134         ElseIf Head = e_Heading.WEST Then
136             MaxX = MinX - 1
138             MinX = MinX - AREA_DIM
140             MaxY = MinY + AREA_DIM * 3 - 1 '+ 26
142             .AreasInfo.MinX = CInt(MinX)
144             .AreasInfo.MinY = CInt(MinY)
        
146         ElseIf Head = e_Heading.EAST Then
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

188                     If MapData(.Pos.Map, X, Y).UserIndex Then Call MakeNPCChar(False, MapData(.Pos.Map, X, Y).UserIndex, NpcIndex, .Pos.Map, .Pos.X, .Pos.Y)

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
208     Call TraceError(Err.Number, Err.Description, "ModAreas.CheckUpdateNeededNpc", Erl)

        
End Sub
 
Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal Map As Integer)
        
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

102         If ConnGroups(Map).UserEntrys(LoopC) = UserIndex Then Exit For
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
122     Call TraceError(Err.Number, Err.Description, "ModAreas.QuitarUser", Erl)

        
End Sub
 
Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal Map As Integer, Optional ByVal appear As Byte = 0)
        
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

106         If ConnGroups(Map).UserEntrys(i) = UserIndex Then
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
       
122         ConnGroups(Map).UserEntrys(TempVal) = UserIndex

        End If
   
        'Update user
124     UserList(UserIndex).AreasInfo.AreaID = 0
   
126     UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
128     UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
130     UserList(UserIndex).AreasInfo.AreaReciveX = 0
132     UserList(UserIndex).AreasInfo.AreaReciveY = 0
   
134     Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, appear)

        
        Exit Sub

AgregarUser_Err:
136     Call TraceError(Err.Number, Err.Description, "ModAreas.AgregarUser", Erl)

        
End Sub
 
Public Sub AgregarNpc(ByVal NpcIndex As Integer)
        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        
        On Error GoTo AgregarNpc_Err
        
100     NpcList(NpcIndex).AreasInfo.AreaID = 0
   
102     NpcList(NpcIndex).AreasInfo.AreaPerteneceX = 0
104     NpcList(NpcIndex).AreasInfo.AreaPerteneceY = 0
106     NpcList(NpcIndex).AreasInfo.AreaReciveX = 0
108     NpcList(NpcIndex).AreasInfo.AreaReciveY = 0
   
110     Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)

        
        Exit Sub

AgregarNpc_Err:
112     Call TraceError(Err.Number, Err.Description, "ModAreas.AgregarNpc", Erl)

        
End Sub

