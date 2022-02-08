Attribute VB_Name = "modSendData"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan  Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget

    ToAll = 1
    ToIndex
    toMap
    ToPCArea
    ToPCAreaButGMs
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToSuperiores
    ToSuperioresArea
    
    ToUsuariosMuertos
    
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ParamArray Args() As Variant)
        
        On Error GoTo SendData_Err
    
        '**************************************************************
        'Author: Juan  Martín Sotuyo Dodero (Maraxus) - Rewrite of original
        'Last Modify Date: 01/08/2007
        'Last modified by: (liquid)
        '**************************************************************
        
        Dim LoopC As Long
        Dim Map   As Integer
        
        Dim Buffer As Network.Writer
        Set Buffer = Protocol_Writes.GetWriterBuffer()
    
100     Select Case sndRoute
            Case SendTarget.ToIndex
                If (UserList(sndIndex).ConnIDValida) Then
                    Call modNetwork.Send(sndIndex, Buffer)
                End If

104         Case SendTarget.ToPCArea
106             Call SendToUserArea(sndIndex, Buffer)

108         Case SendTarget.ToPCAreaButGMs
110             Call SendToUserAreaButGMs(sndIndex, Buffer)

112         Case SendTarget.ToUsuariosMuertos
114             Call SendToUsersMuertosArea(sndIndex, Buffer)

116         Case SendTarget.ToAdmins
118             For LoopC = 1 To LastUser
120                 If UserList(LoopC).ConnIDValida Then
122                     If UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero) Then
124                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
126             Next LoopC

128         Case SendTarget.ToSuperiores
130             For LoopC = 1 To LastUser
132                 If UserList(LoopC).ConnIDValida Then
134                     If CompararPrivilegiosUser(LoopC, sndIndex) > 0 Then
136                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
138             Next LoopC

140         Case SendTarget.ToSuperioresArea
142             Call SendToSuperioresArea(sndIndex, Buffer)

144         Case SendTarget.ToAll
146             For LoopC = 1 To LastUser
148                 If UserList(LoopC).ConnIDValida Then
150                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
152                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
154             Next LoopC

156         Case SendTarget.ToAllButIndex
158             For LoopC = 1 To LastUser
160                 If (UserList(LoopC).ConnIDValida) And (LoopC <> sndIndex) Then
162                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
164                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
166             Next LoopC

168         Case SendTarget.toMap
170             Call SendToMap(sndIndex, Buffer)

172         Case SendTarget.ToMapButIndex
174             Call SendToMapButIndex(sndIndex, Buffer)
        
176         Case SendTarget.ToGuildMembers
178             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

180             While LoopC > 0
182                 If (UserList(LoopC).ConnIDValida) Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                    
186                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
        
188         Case SendTarget.ToDeadArea
190             Call SendToDeadUserArea(sndIndex, Buffer)
        
192         Case SendTarget.ToPCAreaButIndex
194             Call SendToUserAreaButindex(sndIndex, Buffer)
            
196         Case SendTarget.ToAdminAreaButIndex
198             Call SendToAdminAreaButIndex(sndIndex, Buffer)
        
200         Case SendTarget.ToClanArea
202             Call SendToUserGuildArea(sndIndex, Buffer)
        
204         Case SendTarget.ToAdminsAreaButConsejeros
206             Call SendToAdminsButConsejerosArea(sndIndex, Buffer)
        
208         Case SendTarget.ToNPCArea
210             Call SendToNpcArea(sndIndex, Buffer)
        
212         Case SendTarget.ToDiosesYclan
214             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

216             While LoopC > 0
218                 If (UserList(LoopC).ConnIDValida) Then
220                     Call modNetwork.Send(LoopC, Buffer)
                    End If
                    
222                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
            
224             LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

226             While LoopC > 0
228                 If (UserList(LoopC).ConnIDValida) Then
230                     Call modNetwork.Send(LoopC, Buffer)
                    End If

232                 LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
                Wend

234         Case SendTarget.ToConsejo
236             For LoopC = 1 To LastUser
238                 If (UserList(LoopC).ConnIDValida) Then
240                     If UserList(LoopC).flags.Privilegios And e_PlayerType.RoyalCouncil Then
242                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
244             Next LoopC

246         Case SendTarget.ToConsejoCaos
248             For LoopC = 1 To LastUser
250                 If (UserList(LoopC).ConnIDValida) Then
252                     If UserList(LoopC).flags.Privilegios And e_PlayerType.ChaosCouncil Then
254                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
256             Next LoopC

258         Case SendTarget.ToRolesMasters
260             For LoopC = 1 To LastUser
262                 If (UserList(LoopC).ConnIDValida) Then
264                     If UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster Then
266                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
268             Next LoopC

270         Case SendTarget.ToCiudadanos
272             For LoopC = 1 To LastUser
274                 If (UserList(LoopC).ConnIDValida) Then
276                     If Status(LoopC) < 2 Then
278                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
280             Next LoopC

282         Case SendTarget.ToCriminales
284             For LoopC = 1 To LastUser
286                 If (UserList(LoopC).ConnIDValida) Then
288                     If Status(LoopC) = 2 Then
290                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
292             Next LoopC

294         Case SendTarget.ToReal
296             For LoopC = 1 To LastUser
298                 If (UserList(LoopC).ConnIDValida) Then
300                     If UserList(LoopC).Faccion.ArmadaReal = 1 Then
302                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
304             Next LoopC

306         Case SendTarget.ToCaos
308             For LoopC = 1 To LastUser
310                 If (UserList(LoopC).ConnIDValida) Then
312                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
314                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
316             Next LoopC

318         Case SendTarget.ToCiudadanosYRMs
320             For LoopC = 1 To LastUser
322                 If (UserList(LoopC).ConnIDValida) Then
324                     If Status(LoopC) < 2 Or (UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster) <> 0 Then
326                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
328             Next LoopC

330         Case SendTarget.ToCriminalesYRMs
332             For LoopC = 1 To LastUser
334                 If (UserList(LoopC).ConnIDValida) Then
336                     If Status(LoopC) = 2 Or (UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster) <> 0 Then
338                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
340             Next LoopC

342         Case SendTarget.ToRealYRMs
344             For LoopC = 1 To LastUser
346                 If (UserList(LoopC).ConnIDValida) Then
348                     If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster) <> 0 Then
350                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
352             Next LoopC

354         Case SendTarget.ToCaosYRMs
356             For LoopC = 1 To LastUser
358                 If (UserList(LoopC).ConnIDValida) Then
360                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster) <> 0 Then
362                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
364             Next LoopC
        End Select

SendData_Err:
        Call Buffer.Clear
        
        If (Err.Number <> 0) Then
366         Call TraceError(Err.Number, Err.Description, "modSendData.SendData", Erl)
        End If
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserArea_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)

114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

118                 If UserList(tempIndex).ConnIDValida Then

120                     Call modNetwork.Send(tempIndex, Buffer)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToUserArea_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserArea", Erl)

        
End Sub

Private Sub SendToUsersMuertosArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserArea_Err
        

        '**************************************************************
        'Author: Jopi
        'Last Modify Date: 23/06/2021
        'Envio la data a los que estan muertos y a los GMs en el area.
        '**************************************************************
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)

114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
118                 If UserList(tempIndex).ConnIDValida Then
                        
                        ' Envio a los que estan MUERTOS y a los GMs cercanos.

120                     If UserList(tempIndex).flags.Muerto = 1 Or EsGM(tempIndex) Then
                        
122                         Call modNetwork.Send(tempIndex, Buffer)
                            
                        End If

                    End If
                End If
            End If

124     Next LoopC
        
        Exit Sub

SendToUserArea_Err:
126     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUsersMuertosArea", Erl)

        
End Sub

Private Sub SendToSuperioresArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserAreaButindex_Err

        '**************************************************************
        'Author: Jopi
        'Last Modify Date: 27/12/2020
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then

122                 If UserList(tempIndex).ConnIDValida Then
                
124                     If CompararPrivilegiosUser(UserIndex, tempIndex) < 0 Then
126                         Call modNetwork.Send(tempIndex, Buffer)
                        End If

                    End If

                End If

            End If

128     Next LoopC
        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)


        
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserAreaButindex_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY


108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then
122                 If tempIndex <> UserIndex Then
124                     If UserList(tempIndex).ConnIDValida Then
126                         Call modNetwork.Send(tempIndex, Buffer)

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)

        
End Sub

Private Sub SendToAdminAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserAreaButindex_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        'sndData = sndData & ENDC

108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then

122                 If tempIndex <> UserIndex And EsGM(tempIndex) Then

124                     If UserList(tempIndex).ConnIDValida Then
126                         Call modNetwork.Send(tempIndex, Buffer)

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToAdminAreaButIndex", Erl)

        
End Sub

Private Sub SendToUserAreaButGMs(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserAreaButindex_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        
        Dim TempInt   As Integer
        Dim tempIndex As Integer
    
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then

122                 If Not EsGM(tempIndex) Then

124                     If UserList(tempIndex).ConnIDValida Then

126                         Call modNetwork.Send(tempIndex, Buffer)

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)

        
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToDeadUserArea_Err
        
        '**************************************************************
        'Author: Juan  Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

                    'Dead and admins read
118                 If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0) Then
120                     Call modNetwork.Send(tempIndex, Buffer)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToDeadUserArea_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToDeadUserArea", Erl)

        
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToUserGuildArea_Err
        

        '**************************************************************
        'Author: Juan  Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
    
112     For LoopC = 1 To ConnGroups(Map).CountEntrys
114         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
116         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
118             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
120                 If UserList(tempIndex).ConnIDValida And UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Then
122                     Call modNetwork.Send(tempIndex, Buffer)

                    End If

                End If

            End If

124     Next LoopC

        
        Exit Sub

SendToUserGuildArea_Err:
126     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserGuildArea", Erl)

        
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToAdminsButConsejerosArea_Err
        

        '**************************************************************
        'Author: Juan  Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
118                 If UserList(tempIndex).ConnIDValida Then
120                     If UserList(tempIndex).flags.Privilegios And (e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin) Then
                            Call modNetwork.Send(tempIndex, Buffer)
                        End If

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToAdminsButConsejerosArea_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToAdminsButConsejerosArea", Erl)

        
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToNpcArea_Err
        

        '**************************************************************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If NpcIndex = 0 Then Exit Sub
        
102     Map = NpcList(NpcIndex).Pos.Map
104     AreaX = NpcList(NpcIndex).AreasInfo.AreaPerteneceX
106     AreaY = NpcList(NpcIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then
122                 If UserList(tempIndex).ConnIDValida Then
124                     Call modNetwork.Send(tempIndex, Buffer)

                    End If

                End If

            End If

126     Next LoopC

        
        Exit Sub

SendToNpcArea_Err:
128     Call TraceError(Err.Number, Err.Description, "modSendData.SendToNpcArea", Erl)

        
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ParamArray Args() As Variant)
        
        On Error GoTo SendToAreaByPos_Err
        
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        
100     AreaX = 2 ^ (AreaX \ 12)
102     AreaY = 2 ^ (AreaY \ 12)
   
104     If Not MapaValido(Map) Then Exit Sub
      
        Dim Buffer As Network.Writer
        Set Buffer = Protocol_Writes.GetWriterBuffer()
        
106     For LoopC = 1 To ConnGroups(Map).CountEntrys
108         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
           
110         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

112         If TempInt Then  'Esta en el area?
114             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

116             If TempInt Then
118                 If UserList(tempIndex).ConnIDValida Then
120                     Call modNetwork.Send(tempIndex, Buffer)

                    End If

                End If

            End If

122     Next LoopC

SendToAreaByPos_Err:
        
        Call Buffer.Clear
            
        If (Err.Number <> 0) Then
124         Call TraceError(Err.Number, Err.Description, "modSendData.SendToAreaByPos", Erl)
        End If
End Sub

Private Sub SendToMap(ByVal Map As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToMap_Err
        
        '**************************************************************
        'Author: Juan  Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim LoopC     As Long

        Dim tempIndex As Integer
    
100     If Not MapaValido(Map) Then Exit Sub

102     For LoopC = 1 To ConnGroups(Map).CountEntrys
104         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
106         If UserList(tempIndex).ConnIDValida Then
108             Call modNetwork.Send(tempIndex, Buffer)

            End If

110     Next LoopC

        
        Exit Sub

SendToMap_Err:
112     Call TraceError(Err.Number, Err.Description, "modSendData.SendToMap", Erl)

        
End Sub

Private Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
        
        On Error GoTo SendToMapButIndex_Err
        

        '**************************************************************
        'Author: Juan  Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim LoopC     As Long
        Dim Map       As Integer
        Dim tempIndex As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
    
104     If Not MapaValido(Map) Then Exit Sub

106     For LoopC = 1 To ConnGroups(Map).CountEntrys
108         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
110         If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
112             Call modNetwork.Send(tempIndex, Buffer)

            End If

114     Next LoopC

        
        Exit Sub

SendToMapButIndex_Err:
116     Call TraceError(Err.Number, Err.Description, "modSendData.SendToMapButIndex", Erl)

        
End Sub

