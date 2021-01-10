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
            
            Case SendTarget.ToIndex
102             Call EnviarDatosASlot(sndIndex, sndData)
                Exit Sub
            
104         Case SendTarget.ToPCArea
106             Call SendToUserArea(sndIndex, sndData)
                Exit Sub
            
108         Case SendTarget.ToPCAreaButGMs
110             Call SendToUserAreaButGMs(sndIndex, sndData)
                Exit Sub
        
112         Case SendTarget.ToAdmins

114             For LoopC = 1 To LastUser

116                 If UserList(LoopC).ConnID <> -1 Then
118                     If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
120                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

122             Next LoopC

                Exit Sub
            
124         Case SendTarget.ToSuperiores

126             For LoopC = 1 To LastUser

128                 If UserList(LoopC).ConnID <> -1 Then

130                     If CompararPrivilegios(LoopC, sndIndex) > 0 Then
132                         Call EnviarDatosASlot(LoopC, sndData)
                        End If
                    
                    End If

134             Next LoopC

                Exit Sub
            
136         Case SendTarget.ToSuperioresArea
138             Call SendToSuperioresArea(sndIndex, sndData)
                Exit Sub
            
140         Case SendTarget.ToAll

142             For LoopC = 1 To LastUser

144                 If UserList(LoopC).ConnID <> -1 Then
146                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
148                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

150             Next LoopC

                Exit Sub
        
152         Case SendTarget.ToAllButIndex

154             For LoopC = 1 To LastUser

156                 If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
158                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
160                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

162             Next LoopC

                Exit Sub
        
164         Case SendTarget.toMap
166             Call SendToMap(sndIndex, sndData)
                Exit Sub
          
168         Case SendTarget.ToMapButIndex
170             Call SendToMapButIndex(sndIndex, sndData)
                Exit Sub
        
172         Case SendTarget.ToGuildMembers
174             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

176             While LoopC > 0

178                 If (UserList(LoopC).ConnID <> -1) Then
180                     Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)

                    End If

182                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
                Exit Sub
        
184         Case SendTarget.ToDeadArea
186             Call SendToDeadUserArea(sndIndex, sndData)
                Exit Sub
        
188         Case SendTarget.ToPCAreaButIndex
190             Call SendToUserAreaButindex(sndIndex, sndData)
                Exit Sub
            
192         Case SendTarget.ToAdminAreaButIndex
194             Call SendToAdminAreaButIndex(sndIndex, sndData)
                Exit Sub
        
196         Case SendTarget.ToClanArea
198             Call SendToUserGuildArea(sndIndex, sndData)
                Exit Sub
        
200         Case SendTarget.ToAdminsAreaButConsejeros
202             Call SendToAdminsButConsejerosArea(sndIndex, sndData)
                Exit Sub
        
204         Case SendTarget.ToNPCArea
206             Call SendToNpcArea(sndIndex, sndData)
                Exit Sub
        
208         Case SendTarget.ToDiosesYclan
210             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

212             While LoopC > 0

214                 If (UserList(LoopC).ConnID <> -1) Then
216                     Call EnviarDatosASlot(LoopC, sndData)

                    End If

218                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
            
220             LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

222             While LoopC > 0

224                 If (UserList(LoopC).ConnID <> -1) Then
226                     Call EnviarDatosASlot(LoopC, sndData)

                    End If

228                 LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
                Wend
            
                Exit Sub
        
230         Case SendTarget.ToConsejo

232             For LoopC = 1 To LastUser

234                 If (UserList(LoopC).ConnID <> -1) Then
236                     If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
238                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

240             Next LoopC

                Exit Sub
        
242         Case SendTarget.ToConsejoCaos

244             For LoopC = 1 To LastUser

246                 If (UserList(LoopC).ConnID <> -1) Then
248                     If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
250                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

252             Next LoopC

                Exit Sub
        
254         Case SendTarget.ToRolesMasters

256             For LoopC = 1 To LastUser

258                 If (UserList(LoopC).ConnID <> -1) Then
260                     If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
262                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

264             Next LoopC

                Exit Sub
        
266         Case SendTarget.ToCiudadanos

268             For LoopC = 1 To LastUser

270                 If (UserList(LoopC).ConnID <> -1) Then
272                     If Status(LoopC) < 2 Then
274                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

276             Next LoopC

                Exit Sub
        
278         Case SendTarget.ToCriminales

280             For LoopC = 1 To LastUser

282                 If (UserList(LoopC).ConnID <> -1) Then
284                     If Status(LoopC) = 2 Then
286                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

288             Next LoopC

                Exit Sub
        
290         Case SendTarget.ToReal

292             For LoopC = 1 To LastUser

294                 If (UserList(LoopC).ConnID <> -1) Then
296                     If UserList(LoopC).Faccion.ArmadaReal = 1 Then
298                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

300             Next LoopC

                Exit Sub
        
302         Case SendTarget.ToCaos

304             For LoopC = 1 To LastUser

306                 If (UserList(LoopC).ConnID <> -1) Then
308                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
310                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

312             Next LoopC

                Exit Sub
        
314         Case SendTarget.ToCiudadanosYRMs

316             For LoopC = 1 To LastUser

318                 If (UserList(LoopC).ConnID <> -1) Then
320                     If Status(LoopC) < 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
322                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

324             Next LoopC

                Exit Sub
        
326         Case SendTarget.ToCriminalesYRMs

328             For LoopC = 1 To LastUser

330                 If (UserList(LoopC).ConnID <> -1) Then
332                     If Status(LoopC) = 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
334                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

336             Next LoopC

                Exit Sub
        
338         Case SendTarget.ToRealYRMs

340             For LoopC = 1 To LastUser

342                 If (UserList(LoopC).ConnID <> -1) Then
344                     If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
346                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

348             Next LoopC

                Exit Sub
        
350         Case SendTarget.ToCaosYRMs

352             For LoopC = 1 To LastUser

354                 If (UserList(LoopC).ConnID <> -1) Then
356                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
358                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

360             Next LoopC

                Exit Sub

        End Select

        
        Exit Sub

SendData_Err:
362     Call RegistrarError(Err.Number, Err.description, "modSendData.SendData", Erl)

        
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

