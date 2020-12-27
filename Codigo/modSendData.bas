Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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
    
End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)
        
        On Error GoTo SendData_Err
    
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
        'Last Modify Date: 01/08/2007
        'Last modified by: (liquid)
        '**************************************************************
        
        Dim LoopC As Long
        Dim Map   As Integer

100     Select Case sndRoute

            Case SendTarget.ToPCArea
102             Call SendToUserArea(sndIndex, sndData)
                Exit Sub
            
104         Case SendTarget.ToPCAreaButGMs
106             Call SendToUserAreaButGMs(sndIndex, sndData)
                Exit Sub
        
108         Case SendTarget.ToAdmins

110             For LoopC = 1 To LastUser

112                 If UserList(LoopC).ConnID <> -1 Then
114                     If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
116                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

118             Next LoopC

                Exit Sub
            
120         Case SendTarget.ToSuperiores

122             For LoopC = 1 To LastUser

124                 If UserList(LoopC).ConnID <> -1 Then

126                     If CompararPrivilegios(LoopC, sndIndex) > 0 Then
128                         Call EnviarDatosASlot(LoopC, sndData)
                        End If
                    
                    End If

130             Next LoopC

                Exit Sub
            
132         Case SendTarget.ToSuperioresArea
134             Call SendToSuperioresArea(sndIndex, sndData)
                Exit Sub
            
136         Case SendTarget.ToAll

138             For LoopC = 1 To LastUser

140                 If UserList(LoopC).ConnID <> -1 Then
142                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
144                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

146             Next LoopC

                Exit Sub
        
148         Case SendTarget.ToAllButIndex

150             For LoopC = 1 To LastUser

152                 If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
154                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
156                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

158             Next LoopC

                Exit Sub
        
160         Case SendTarget.toMap
162             Call SendToMap(sndIndex, sndData)
                Exit Sub
          
164         Case SendTarget.ToMapButIndex
166             Call SendToMapButIndex(sndIndex, sndData)
                Exit Sub
        
168         Case SendTarget.ToGuildMembers
170             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

172             While LoopC > 0

174                 If (UserList(LoopC).ConnID <> -1) Then
176                     Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)

                    End If

178                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
                Exit Sub
        
180         Case SendTarget.ToDeadArea
182             Call SendToDeadUserArea(sndIndex, sndData)
                Exit Sub
        
184         Case SendTarget.ToPCAreaButIndex
186             Call SendToUserAreaButindex(sndIndex, sndData)
                Exit Sub
            
188         Case SendTarget.ToAdminAreaButIndex
190             Call SendToAdminAreaButIndex(sndIndex, sndData)
                Exit Sub
        
192         Case SendTarget.ToClanArea
194             Call SendToUserGuildArea(sndIndex, sndData)
                Exit Sub
        
196         Case SendTarget.ToAdminsAreaButConsejeros
198             Call SendToAdminsButConsejerosArea(sndIndex, sndData)
                Exit Sub
        
200         Case SendTarget.ToNPCArea
202             Call SendToNpcArea(sndIndex, sndData)
                Exit Sub
        
204         Case SendTarget.ToDiosesYclan
206             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

208             While LoopC > 0

210                 If (UserList(LoopC).ConnID <> -1) Then
212                     Call EnviarDatosASlot(LoopC, sndData)

                    End If

214                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
            
216             LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

218             While LoopC > 0

220                 If (UserList(LoopC).ConnID <> -1) Then
222                     Call EnviarDatosASlot(LoopC, sndData)

                    End If

224                 LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
                Wend
            
                Exit Sub
        
226         Case SendTarget.ToConsejo

228             For LoopC = 1 To LastUser

230                 If (UserList(LoopC).ConnID <> -1) Then
232                     If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
234                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

236             Next LoopC

                Exit Sub
        
238         Case SendTarget.ToConsejoCaos

240             For LoopC = 1 To LastUser

242                 If (UserList(LoopC).ConnID <> -1) Then
244                     If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
246                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

248             Next LoopC

                Exit Sub
        
250         Case SendTarget.ToRolesMasters

252             For LoopC = 1 To LastUser

254                 If (UserList(LoopC).ConnID <> -1) Then
256                     If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
258                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

260             Next LoopC

                Exit Sub
        
262         Case SendTarget.ToCiudadanos

264             For LoopC = 1 To LastUser

266                 If (UserList(LoopC).ConnID <> -1) Then
268                     If Status(LoopC) < 2 Then
270                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

272             Next LoopC

                Exit Sub
        
274         Case SendTarget.ToCriminales

276             For LoopC = 1 To LastUser

278                 If (UserList(LoopC).ConnID <> -1) Then
280                     If Status(LoopC) = 2 Then
282                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

284             Next LoopC

                Exit Sub
        
286         Case SendTarget.ToReal

288             For LoopC = 1 To LastUser

290                 If (UserList(LoopC).ConnID <> -1) Then
292                     If UserList(LoopC).Faccion.ArmadaReal = 1 Then
294                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

296             Next LoopC

                Exit Sub
        
298         Case SendTarget.ToCaos

300             For LoopC = 1 To LastUser

302                 If (UserList(LoopC).ConnID <> -1) Then
304                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
306                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

308             Next LoopC

                Exit Sub
        
310         Case SendTarget.ToCiudadanosYRMs

312             For LoopC = 1 To LastUser

314                 If (UserList(LoopC).ConnID <> -1) Then
316                     If Status(LoopC) < 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
318                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

320             Next LoopC

                Exit Sub
        
322         Case SendTarget.ToCriminalesYRMs

324             For LoopC = 1 To LastUser

326                 If (UserList(LoopC).ConnID <> -1) Then
328                     If Status(LoopC) = 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
330                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

332             Next LoopC

                Exit Sub
        
334         Case SendTarget.ToRealYRMs

336             For LoopC = 1 To LastUser

338                 If (UserList(LoopC).ConnID <> -1) Then
340                     If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
342                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

344             Next LoopC

                Exit Sub
        
346         Case SendTarget.ToCaosYRMs

348             For LoopC = 1 To LastUser

350                 If (UserList(LoopC).ConnID <> -1) Then
352                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
354                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

356             Next LoopC

                Exit Sub

        End Select

        
        Exit Sub

SendData_Err:
358     Call RegistrarError(Err.Number, Err.description, "modSendData.SendData", Erl)

        
End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sndData As String)
        
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
120                     Call EnviarDatosASlot(tempIndex, sndData)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToUserArea_Err:
124     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserArea", Erl)
126     Resume Next
        
End Sub

Private Sub SendToSuperioresArea(ByVal UserIndex As Integer, ByVal sndData As String)
        
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
                
124                     If CompararPrivilegios(UserIndex, tempIndex) < 0 Then
126                         Call EnviarDatosASlot(tempIndex, sndData)
                        End If

                    End If

                End If

            End If

128     Next LoopC
        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserAreaButindex", Erl)

132     Resume Next
        
End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sndData As String)
        
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
126                         Call EnviarDatosASlot(tempIndex, sndData)

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserAreaButindex", Erl)
132     Resume Next
        
End Sub

Private Sub SendToAdminAreaButIndex(ByVal UserIndex As Integer, ByVal sndData As String)
        
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
126                         Call EnviarDatosASlot(tempIndex, sndData)

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToAdminAreaButIndex", Erl)
132     Resume Next
        
End Sub

Private Sub SendToUserAreaButGMs(ByVal UserIndex As Integer, ByVal sndData As String)
        
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

126                         Call EnviarDatosASlot(tempIndex, sndData)

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserAreaButindex", Erl)
132     Resume Next
        
End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sndData As String)
        
        On Error GoTo SendToDeadUserArea_Err
        
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
118                 If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
120                     Call EnviarDatosASlot(tempIndex, sndData)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToDeadUserArea_Err:
124     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToDeadUserArea", Erl)
126     Resume Next
        
End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sndData As String)
        
        On Error GoTo SendToUserGuildArea_Err
        

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
122                     Call EnviarDatosASlot(tempIndex, sndData)

                    End If

                End If

            End If

124     Next LoopC

        
        Exit Sub

SendToUserGuildArea_Err:
126     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserGuildArea", Erl)
128     Resume Next
        
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sndData As String)
        
        On Error GoTo SendToAdminsButConsejerosArea_Err
        

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
120                     If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then Call EnviarDatosASlot(tempIndex, sndData)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToAdminsButConsejerosArea_Err:
124     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToAdminsButConsejerosArea", Erl)
126     Resume Next
        
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sndData As String)
        
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
        
102     Map = Npclist(NpcIndex).Pos.Map
104     AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
106     AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then
122                 If UserList(tempIndex).ConnIDValida Then
124                     Call EnviarDatosASlot(tempIndex, sndData)

                    End If

                End If

            End If

126     Next LoopC

        
        Exit Sub

SendToNpcArea_Err:
128     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToNpcArea", Erl)
130     Resume Next
        
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sndData As String)
        
        On Error GoTo SendToAreaByPos_Err
        
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        
100     AreaX = 2 ^ (AreaX \ 12)
102     AreaY = 2 ^ (AreaY \ 12)
   
104     If Not MapaValido(Map) Then Exit Sub
 
106     For LoopC = 1 To ConnGroups(Map).CountEntrys
108         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
           
110         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

112         If TempInt Then  'Esta en el area?
114             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

116             If TempInt Then
118                 If UserList(tempIndex).ConnIDValida Then
120                     Call EnviarDatosASlot(tempIndex, sndData)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToAreaByPos_Err:
124     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToAreaByPos", Erl)
126     Resume Next
        
End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sndData As String)
        
        On Error GoTo SendToMap_Err
        
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim LoopC     As Long

        Dim tempIndex As Integer
    
100     If Not MapaValido(Map) Then Exit Sub

102     For LoopC = 1 To ConnGroups(Map).CountEntrys
104         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
106         If UserList(tempIndex).ConnIDValida Then
108             Call EnviarDatosASlot(tempIndex, sndData)

            End If

110     Next LoopC

        
        Exit Sub

SendToMap_Err:
112     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToMap", Erl)
114     Resume Next
        
End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sndData As String)
        
        On Error GoTo SendToMapButIndex_Err
        

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
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
112             Call EnviarDatosASlot(tempIndex, sndData)

            End If

114     Next LoopC

        
        Exit Sub

SendToMapButIndex_Err:
116     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToMapButIndex", Erl)
118     Resume Next
        
End Sub

