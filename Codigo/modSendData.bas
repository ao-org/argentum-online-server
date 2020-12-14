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

End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndData As String)

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus) - Rewrite of original
        'Last Modify Date: 01/08/2007
        'Last modified by: (liquid)
        '**************************************************************
        On Error Resume Next

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
126                     If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
128                         If UserList(LoopC).flags.Privilegios >= PlayerType.Admin Then
130                             Call EnviarDatosASlot(LoopC, sndData)

                            End If

                        End If

                    End If

132             Next LoopC

                Exit Sub
            
134         Case SendTarget.ToAll

136             For LoopC = 1 To LastUser

138                 If UserList(LoopC).ConnID <> -1 Then
140                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
142                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

144             Next LoopC

                Exit Sub
        
146         Case SendTarget.ToAllButIndex

148             For LoopC = 1 To LastUser

150                 If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
152                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
154                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

156             Next LoopC

                Exit Sub
        
158         Case SendTarget.toMap
160             Call SendToMap(sndIndex, sndData)
                Exit Sub
          
162         Case SendTarget.ToMapButIndex
164             Call SendToMapButIndex(sndIndex, sndData)
                Exit Sub
        
166         Case SendTarget.ToGuildMembers
168             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

170             While LoopC > 0

172                 If (UserList(LoopC).ConnID <> -1) Then
174                     Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)

                    End If

176                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
                Exit Sub
        
178         Case SendTarget.ToDeadArea
180             Call SendToDeadUserArea(sndIndex, sndData)
                Exit Sub
        
182         Case SendTarget.ToPCAreaButIndex
184             Call SendToUserAreaButindex(sndIndex, sndData)
                Exit Sub
            
186         Case SendTarget.ToAdminAreaButIndex
188             Call SendToAdminAreaButIndex(sndIndex, sndData)
                Exit Sub
        
190         Case SendTarget.ToClanArea
192             Call SendToUserGuildArea(sndIndex, sndData)
                Exit Sub
        
194         Case SendTarget.ToAdminsAreaButConsejeros
196             Call SendToAdminsButConsejerosArea(sndIndex, sndData)
                Exit Sub
        
198         Case SendTarget.ToNPCArea
200             Call SendToNpcArea(sndIndex, sndData)
                Exit Sub
        
202         Case SendTarget.ToDiosesYclan
204             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

206             While LoopC > 0

208                 If (UserList(LoopC).ConnID <> -1) Then
210                     Call EnviarDatosASlot(LoopC, sndData)

                    End If

212                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
            
214             LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

216             While LoopC > 0

218                 If (UserList(LoopC).ConnID <> -1) Then
220                     Call EnviarDatosASlot(LoopC, sndData)

                    End If

222                 LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
                Wend
            
                Exit Sub
        
224         Case SendTarget.ToConsejo

226             For LoopC = 1 To LastUser

228                 If (UserList(LoopC).ConnID <> -1) Then
230                     If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
232                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

234             Next LoopC

                Exit Sub
        
236         Case SendTarget.ToConsejoCaos

238             For LoopC = 1 To LastUser

240                 If (UserList(LoopC).ConnID <> -1) Then
242                     If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
244                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

246             Next LoopC

                Exit Sub
        
248         Case SendTarget.ToRolesMasters

250             For LoopC = 1 To LastUser

252                 If (UserList(LoopC).ConnID <> -1) Then
254                     If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
256                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

258             Next LoopC

                Exit Sub
        
260         Case SendTarget.ToCiudadanos

262             For LoopC = 1 To LastUser

264                 If (UserList(LoopC).ConnID <> -1) Then
266                     If Status(LoopC) < 2 Then
268                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

270             Next LoopC

                Exit Sub
        
272         Case SendTarget.ToCriminales

274             For LoopC = 1 To LastUser

276                 If (UserList(LoopC).ConnID <> -1) Then
278                     If Status(LoopC) = 2 Then
280                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

282             Next LoopC

                Exit Sub
        
284         Case SendTarget.ToReal

286             For LoopC = 1 To LastUser

288                 If (UserList(LoopC).ConnID <> -1) Then
290                     If UserList(LoopC).Faccion.ArmadaReal = 1 Then
292                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

294             Next LoopC

                Exit Sub
        
296         Case SendTarget.ToCaos

298             For LoopC = 1 To LastUser

300                 If (UserList(LoopC).ConnID <> -1) Then
302                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
304                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

306             Next LoopC

                Exit Sub
        
308         Case SendTarget.ToCiudadanosYRMs

310             For LoopC = 1 To LastUser

312                 If (UserList(LoopC).ConnID <> -1) Then
314                     If Status(LoopC) < 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
316                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

318             Next LoopC

                Exit Sub
        
320         Case SendTarget.ToCriminalesYRMs

322             For LoopC = 1 To LastUser

324                 If (UserList(LoopC).ConnID <> -1) Then
326                     If Status(LoopC) = 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
328                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

330             Next LoopC

                Exit Sub
        
332         Case SendTarget.ToRealYRMs

334             For LoopC = 1 To LastUser

336                 If (UserList(LoopC).ConnID <> -1) Then
338                     If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
340                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

342             Next LoopC

                Exit Sub
        
344         Case SendTarget.ToCaosYRMs

346             For LoopC = 1 To LastUser

348                 If (UserList(LoopC).ConnID <> -1) Then
350                     If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
352                         Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

354             Next LoopC

                Exit Sub

        End Select

End Sub

Private Sub SendToUserArea(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY
    
106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)

112         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
114             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
116                 If UserList(tempIndex).ConnIDValida Then
118                     Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

120     Next LoopC

        
        Exit Sub

SendToUserArea_Err:
122     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserArea", Erl)
124     Resume Next
        
End Sub

Private Sub SendToUserAreaButindex(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY
        'sdData = sdData & ENDC

106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
112         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

114         If TempInt Then  'Esta en el area?
116             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

118             If TempInt Then
120                 If tempIndex <> Userindex Then
122                     If UserList(tempIndex).ConnIDValida Then
124                         Call EnviarDatosASlot(tempIndex, sdData)

                        End If

                    End If

                End If

            End If

126     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
128     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserAreaButindex", Erl)
130     Resume Next
        
End Sub

Private Sub SendToAdminAreaButIndex(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY
        'sdData = sdData & ENDC

106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
112         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

114         If TempInt Then  'Esta en el area?
116             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

118             If TempInt Then

120                 If tempIndex <> Userindex And EsGM(tempIndex) Then

122                     If UserList(tempIndex).ConnIDValida Then
124                         Call EnviarDatosASlot(tempIndex, sdData)

                        End If

                    End If

                End If

            End If

126     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
128     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToAdminAreaButIndex", Erl)
130     Resume Next
        
End Sub

Private Sub SendToUserAreaButGMs(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY

106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
112         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

114         If TempInt Then  'Esta en el area?
116             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

118             If TempInt Then

120                 If Not EsGM(tempIndex) Then

122                     If UserList(tempIndex).ConnIDValida Then

124                         Call EnviarDatosASlot(tempIndex, sdData)

                        End If

                    End If

                End If

            End If

126     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
128     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserAreaButindex", Erl)
130     Resume Next
        
End Sub

Private Sub SendToDeadUserArea(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY
    
106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
112         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
114             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

                    'Dead and admins read
116                 If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
118                     Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

120     Next LoopC

        
        Exit Sub

SendToDeadUserArea_Err:
122     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToDeadUserArea", Erl)
124     Resume Next
        
End Sub

Private Sub SendToUserGuildArea(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY
    
106     If Not MapaValido(Map) Then Exit Sub
    
108     If UserList(Userindex).GuildIndex = 0 Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
118                 If UserList(tempIndex).ConnIDValida And UserList(tempIndex).GuildIndex = UserList(Userindex).GuildIndex Then
120                     Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToUserGuildArea_Err:
124     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToUserGuildArea", Erl)
126     Resume Next
        
End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal Userindex As Integer, ByVal sdData As String)
        
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
    
100     Map = UserList(Userindex).Pos.Map
102     AreaX = UserList(Userindex).AreasInfo.AreaPerteneceX
104     AreaY = UserList(Userindex).AreasInfo.AreaPerteneceY
    
106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
112         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
114             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
116                 If UserList(tempIndex).ConnIDValida Then
118                     If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

120     Next LoopC

        
        Exit Sub

SendToAdminsButConsejerosArea_Err:
122     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToAdminsButConsejerosArea", Erl)
124     Resume Next
        
End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)
        
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
    
100     Map = Npclist(NpcIndex).Pos.Map
102     AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
104     AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
        'sdData = sdData & ENDC
    
106     If Not MapaValido(Map) Then Exit Sub
    
108     For LoopC = 1 To ConnGroups(Map).CountEntrys
110         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
112         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

114         If TempInt Then  'Esta en el area?
116             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

118             If TempInt Then
120                 If UserList(tempIndex).ConnIDValida Then
122                     Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

124     Next LoopC

        
        Exit Sub

SendToNpcArea_Err:
126     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToNpcArea", Erl)
128     Resume Next
        
End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
        
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
120                     Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToAreaByPos_Err:
124     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToAreaByPos", Erl)
126     Resume Next
        
End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)
        
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
108             Call EnviarDatosASlot(tempIndex, sdData)

            End If

110     Next LoopC

        
        Exit Sub

SendToMap_Err:
112     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToMap", Erl)
114     Resume Next
        
End Sub

Public Sub SendToMapButIndex(ByVal Userindex As Integer, ByVal sdData As String)
        
        On Error GoTo SendToMapButIndex_Err
        

        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 5/24/2007
        '
        '**************************************************************
        Dim LoopC     As Long

        Dim Map       As Integer

        Dim tempIndex As Integer
    
100     Map = UserList(Userindex).Pos.Map
    
102     If Not MapaValido(Map) Then Exit Sub

104     For LoopC = 1 To ConnGroups(Map).CountEntrys
106         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
108         If tempIndex <> Userindex And UserList(tempIndex).ConnIDValida Then
110             Call EnviarDatosASlot(tempIndex, sdData)

            End If

112     Next LoopC

        
        Exit Sub

SendToMapButIndex_Err:
114     Call RegistrarError(Err.Number, Err.description, "modSendData.SendToMapButIndex", Erl)
116     Resume Next
        
End Sub

