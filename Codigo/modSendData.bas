Attribute VB_Name = "modSendData"
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

Public Enum SendTarget
    ToAll = 1
    ToIndex
    toMap
    ToPCArea
    ToPCAliveArea
    ToPCAreaButGMs
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToNPCAliveArea
    ToNPCDeadArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToPCAliveAreaButIndex
    ToAdminAreaButIndex
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToSuperiores
    ToSuperioresArea
    ToPCDeadArea
    ToPCDeadAreaButIndex
    ToAdminsYDioses
    ToJugadoresCaptura
    ToPCAreaButFollowerAndIndex
    ToGroup
    ToGroupButIndex
End Enum

Public Sub SendToConnection(ByVal ConnectionId, Optional Args As Variant)
On Error GoTo SendToConnection_Err
#If DIRECT_PLAY = 0 Then
    Dim writer As Network.writer
    Set writer = Protocol_Writes.GetWriterBuffer()
#Else
    Dim writer As clsNetWriter
    Set writer = Protocol_Writes.writer
#End If

    Call modNetwork.SendToConnection(ConnectionID, writer)
    writer.Clear
    Exit Sub
SendToConnection_Err:
    Call TraceError(Err.Number, Err.Description, "modSendData.SendToConnection", Erl)
    Call writer.Clear
End Sub

Public Sub SendData(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, Optional Args As Variant, Optional ByVal validateInvi As Boolean = False)
On Error GoTo SendData_Err
    
#If DIRECT_PLAY = 0 Then
     Dim buffer As Network.writer
        Set Buffer = Protocol_Writes.GetWriterBuffer()
#Else
    Dim buffer As clsNetWriter
    Set buffer = Protocol_Writes.writer
#End If
    
   Dim LoopC As Long
   Dim Map   As Integer
    
   Select Case sndRoute
            Case SendTarget.ToIndex
                Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
                With UserList(sndIndex)
                    If (.ConnectionDetails.ConnIDValida) Then
                        Call modNetwork.Send(sndIndex, buffer)
                    End If
                End With

104         Case SendTarget.ToPCArea
                Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
106             Call SendToUserArea(sndIndex, Buffer, validateInvi)

109         Case SendTarget.ToPCAliveArea
111             Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
                Call SendToUserAliveArea(sndIndex, Buffer, ValidateInvi)

            
105         Case SendTarget.ToPCAreaButFollowerAndIndex
                Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
107             Call SendToUserAreaButFollowerAndIndex(sndIndex, Buffer)

108         Case SendTarget.ToPCAreaButGMs
                Debug.Assert sndIndex >= LBound(UserList) And sndIndex <= UBound(UserList)
110             Call SendToUserAreaButGMs(sndIndex, Buffer)

112         Case SendTarget.ToPCDeadArea
114             Call SendToPCDeadArea(sndIndex, Buffer)
            
            Case SendTarget.ToPCDeadAreaButIndex
                Call SendToPCDeadAreaButIndex(sndIndex, Buffer)

116         Case SendTarget.ToAdmins
118             For LoopC = 1 To LastUser
120                 If UserList(LoopC).ConnectionDetails.ConnIDValida Then
122                     If UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero) Then
124                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
126             Next LoopC
            
            Case SendTarget.ToAdminsYDioses
                For LoopC = 1 To LastUser
                    If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                        If UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios) Then
                            Call modNetwork.Send(LoopC, Buffer)
                           End If
                       End If
                Next LoopC
                
                
            Case SendTarget.ToJugadoresCaptura
                For LoopC = 1 To LastUser
                    If UserList(LoopC).ConnectionDetails.ConnIDValida Then
                        If UserList(LoopC).flags.jugando_captura = 1 Then
                            Call modNetwork.Send(LoopC, Buffer)
                           End If
                       End If
                Next LoopC


128         Case SendTarget.ToSuperiores
130             For LoopC = 1 To LastUser
132                 If UserList(LoopC).ConnectionDetails.ConnIDValida Then
134                     If CompararPrivilegiosUser(LoopC, sndIndex) > 0 Then
136                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
138             Next LoopC

140         Case SendTarget.ToSuperioresArea
142             Call SendToSuperioresArea(sndIndex, Buffer)

144         Case SendTarget.ToAll
146             For LoopC = 1 To LastUser
148                 If UserList(LoopC).ConnectionDetails.ConnIDValida Then
150                     If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
152                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
154             Next LoopC

156         Case SendTarget.ToAllButIndex
158             For LoopC = 1 To LastUser
160                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) And (LoopC <> sndIndex) Then
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
182                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                        Call modNetwork.Send(LoopC, Buffer)
                    End If
                    
186                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
                
192         Case SendTarget.ToPCAreaButIndex
194             Call SendToUserAreaButindex(sndIndex, Buffer, validateInvi)

193         Case SendTarget.ToPCAliveAreaButIndex
195             Call SendToUserAliveAreaButindex(sndIndex, Buffer, validateInvi)

196         Case SendTarget.ToAdminAreaButIndex
198             Call SendToAdminAreaButIndex(sndIndex, Buffer)
        
200         Case SendTarget.ToClanArea
202             Call SendToUserGuildArea(sndIndex, Buffer)
                
208         Case SendTarget.ToNPCArea
210             Call SendToNpcArea(sndIndex, Buffer)
        
209         Case SendTarget.ToNPCAliveArea
211             Call SendToNpcAliveArea(sndIndex, Buffer)

212         Case SendTarget.ToDiosesYclan
214             LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

216             While LoopC > 0
218                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
220                     Call modNetwork.Send(LoopC, Buffer)
                    End If
                    
222                 LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
                Wend
            
224             LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

226             While LoopC > 0
228                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
230                     Call modNetwork.Send(LoopC, Buffer)
                    End If

232                 LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
                Wend

234         Case SendTarget.ToConsejo
236             For LoopC = 1 To LastUser
238                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
240                     If UserList(LoopC).Faccion.Status = e_Facciones.consejo Then
242                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
244             Next LoopC

246         Case SendTarget.ToConsejoCaos
248             For LoopC = 1 To LastUser
250                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
                        If UserList(LoopC).Faccion.Status = e_Facciones.concilio Then
254                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
256             Next LoopC

258         Case SendTarget.ToRolesMasters
260             For LoopC = 1 To LastUser
262                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
264                     If UserList(LoopC).flags.Privilegios And e_PlayerType.RoleMaster Then
266                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
268             Next LoopC

342         Case SendTarget.ToRealYRMs
344             For LoopC = 1 To LastUser
346                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
348                     If UserList(LoopC).Faccion.Status = e_Facciones.Armada Or _
                        (UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0 Or _
                        UserList(LoopC).Faccion.Status = e_Facciones.consejo Then
350                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
352             Next LoopC

354         Case SendTarget.ToCaosYRMs
356             For LoopC = 1 To LastUser
358                 If (UserList(LoopC).ConnectionDetails.ConnIDValida) Then
360                     If UserList(LoopC).Faccion.Status = e_Facciones.Caos Or _
                        (UserList(LoopC).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0 Or _
                        UserList(LoopC).Faccion.Status = e_Facciones.concilio Then
362                         Call modNetwork.Send(LoopC, Buffer)
                        End If
                    End If
364             Next LoopC
            Case SendTarget.ToGroup
                Call SendToGroup(sndIndex, Buffer)
            Case SendTarget.ToGroupButIndex
                Call SendToGroupButIndex(sndIndex, Buffer)
        End Select

SendData_Err:
        Call Buffer.Clear
        
        If (Err.Number <> 0) Then
366         Call TraceError(Err.Number, Err.Description, "modSendData.SendData", Erl)
        End If
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserAliveArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer, Optional ByVal validateInvi As Boolean = False)
#Else
Private Sub SendToUserAliveArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean = False)
#End If
On Error GoTo SendToUserArea_Err

        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        Dim enviaDatos As Boolean
        
100     If UserIndex = 0 Then Exit Sub
        
102     map = UserList(UserIndex).Pos.map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(map).CountEntrys
112         tempIndex = ConnGroups(map).UserEntrys(LoopC)

114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
118                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        If UserList(tempIndex).flags.Muerto = 0 Or MapInfo(UserList(tempIndex).pos.Map).Seguro = 1 Or (UserList(UserIndex).GuildIndex > 0 And UserList(UserIndex).GuildIndex = UserList(tempIndex).GuildIndex) Or IsSet(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTalkToDead) Then
                            enviaDatos = True
                            If Not EsGM(tempIndex) Then
                                If UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0 And validateInvi And Not (UserList(tempIndex).GuildIndex > 0 And UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex And modGuilds.NivelDeClan(UserList(tempIndex).GuildIndex) >= 6) And UserList(UserIndex).flags.Navegando = 0 Then
                                    If Distancia(UserList(UserIndex).Pos, UserList(tempIndex).Pos) > DISTANCIA_ENVIO_DATOS And UserList(UserIndex).Counters.timeFx + UserList(UserIndex).Counters.timeChat = 0 Then
                                        enviaDatos = False
                                    End If
                                End If
                            End If
                            
                            If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                                    Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                            End If
                            
                            If enviaDatos Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If
                            
                        End If
                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToUserArea_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserArea", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer, Optional ByVal validateInvi As Boolean)
#Else
Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean)
#End If
        
        On Error GoTo SendToUserArea_Err

        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        Dim enviaDatos As Boolean
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)

114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

118                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        enviaDatos = True
                        If Not EsGM(tempIndex) Then
                            If UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0 And validateInvi Then
                                If Distancia(UserList(UserIndex).Pos, UserList(tempIndex).Pos) > DISTANCIA_ENVIO_DATOS And UserList(UserIndex).Counters.timeFx + UserList(UserIndex).Counters.timeChat = 0 Then
                                    enviaDatos = False
                                End If
                            End If
                        End If
                        
                        If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                                Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                        End If
                        
                        If enviaDatos Then
                            Call modNetwork.Send(tempIndex, Buffer)
                        End If

                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToUserArea_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserArea", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserAreaButFollowerAndIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToUserAreaButFollowerAndIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToUserAreaButFollower_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     map = UserList(UserIndex).Pos.map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(map) Then Exit Sub
        
110     For LoopC = 1 To ConnGroups(map).CountEntrys
112         tempIndex = ConnGroups(map).UserEntrys(LoopC)

114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

118                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        If UserList(tempIndex).flags.SigueUsuario.ArrayIndex = 0 And tempIndex <> userIndex Then
120                         Call modNetwork.Send(tempIndex, Buffer)
                        End If
                    End If

                End If

            End If

122     Next LoopC

        
        Exit Sub

SendToUserAreaButFollower_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButFollower", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToPCDeadArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToPCDeadArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
 On Error GoTo SendToUserArea_Err
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
118                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        
                        ' Envio a los que estan MUERTOS y a los GMs cercanos.

120                     If UserList(tempIndex).flags.Muerto = 1 Or EsGM(tempIndex) Or IsSet(UserList(tempIndex).flags.StatusMask, e_StatusMask.eTalkToDead) Then
                        
122                         Call modNetwork.Send(tempIndex, Buffer)
                            
                        End If

                    End If
                End If
            End If

124     Next LoopC
        
        Exit Sub

SendToUserArea_Err:
126     Call TraceError(Err.Number, Err.Description, "modSendData.SendToPCDeadArea", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToPCDeadAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToPCDeadAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
        
On Error GoTo SendToUserArea_Err

        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).pos.Map
104     AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
106     AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        
108     If Not MapaValido(Map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(Map).CountEntrys
112         tempIndex = ConnGroups(Map).UserEntrys(LoopC)

114         If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
116             If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
118                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        ' Envio a los que estan MUERTOS y a los GMs cercanos.
                        If tempIndex <> UserIndex Then
120                         If UserList(tempIndex).flags.Muerto = 1 Or IsSet(UserList(tempIndex).flags.StatusMask, e_StatusMask.eTalkToDead) Then
122                             Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If
                    End If
                End If
            End If

124     Next LoopC
        
        Exit Sub

SendToUserArea_Err:
126     Call TraceError(Err.Number, Err.Description, "modSendData.SendToPCDeadArea", Erl)

        
End Sub


#If DIRECT_PLAY = 0 Then
Private Sub SendToSuperioresArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToSuperioresArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToSuperioresArea_Err
        
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

122                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                
124                     If CompararPrivilegiosUser(UserIndex, tempIndex) < 0 Then
126                         Call modNetwork.Send(tempIndex, Buffer)
                        End If

                    End If

                End If

            End If

128     Next LoopC
        
        Exit Sub

SendToSuperioresArea_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToSuperioresArea", Erl)


        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer, Optional ByVal validateInvi As Boolean = False)
#Else
Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean = False)
#End If
        
On Error GoTo SendToUserAreaButindex_Err
        

        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        Dim enviaDatos As Boolean
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
124                     If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
126                         enviaDatos = True
                            
                            If Not EsGM(tempIndex) Then
                                If UserList(UserIndex).flags.invisible + UserList(UserIndex).flags.Oculto > 0 And validateInvi Then
                                    If Distancia(UserList(UserIndex).Pos, UserList(tempIndex).Pos) > DISTANCIA_ENVIO_DATOS And UserList(UserIndex).Counters.timeFx + UserList(UserIndex).Counters.timeChat = 0 Then
                                        enviaDatos = False
                                    End If
                                End If
                            End If
                            
                            If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                                    Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                            End If
                            
                            If enviaDatos Then
                                Call modNetwork.Send(tempIndex, Buffer)
                            End If

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Function CanSendToUser(ByRef SourceUser As t_User, ByRef TargetUser As t_User, ByVal TargetIndex As Integer, ByRef Buffer As Network.Writer, ByVal ValidateInvi As Boolean) As Boolean
#Else
Private Function CanSendToUser(ByRef SourceUser As t_User, ByRef TargetUser As t_User, ByVal TargetIndex As Integer, ByRef Buffer As clsNetWriter, ByVal ValidateInvi As Boolean) As Boolean
#End If
    If (TargetUser.AreasInfo.AreaReciveX And SourceUser.AreasInfo.AreaPerteneceX) = 0 Then Exit Function
    If (TargetUser.AreasInfo.AreaReciveY And SourceUser.AreasInfo.AreaPerteneceY) = 0 Then Exit Function
    If Not TargetUser.ConnectionDetails.ConnIDValida Then Exit Function
    If Not (TargetUser.flags.Muerto = 0 Or MapInfo(TargetUser.pos.Map).Seguro = 1 Or (SourceUser.GuildIndex > 0 And SourceUser.GuildIndex = TargetUser.GuildIndex) Or IsSet(TargetUser.flags.StatusMask, e_StatusMask.eTalkToDead) Or IsSet(SourceUser.flags.StatusMask, e_StatusMask.eTalkToDead)) Then Exit Function
    If IsValidUserRef(TargetUser.flags.GMMeSigue) Then
        Call modNetwork.Send(TargetUser.flags.GMMeSigue.ArrayIndex, Buffer)
    End If
    If Not EsGM(TargetIndex) Then
        If SourceUser.flags.invisible + SourceUser.flags.Oculto > 0 And ValidateInvi And Not CheckGuildSend(SourceUser, TargetUser) And SourceUser.flags.Navegando = 0 Then
            If Distancia(SourceUser.pos, TargetUser.pos) > DISTANCIA_ENVIO_DATOS And SourceUser.Counters.timeFx + SourceUser.Counters.timeChat = 0 Then
                Exit Function
            End If
        End If
    End If
    CanSendToUser = True
End Function

Public Function CheckGuildSend(ByRef SourceUser As t_User, ByRef TargetUser As t_User) As Boolean
    CheckGuildSend = False
    If SourceUser.GuildIndex = 0 Then Exit Function
    If SourceUser.GuildIndex <> TargetUser.GuildIndex Then Exit Function
    If modGuilds.NivelDeClan(TargetUser.GuildIndex) < 6 Then
        CheckGuildSend = SourceUser.Counters.timeGuildChat > 0
        Exit Function
    End If
    CheckGuildSend = True
End Function

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserAliveAreaButindex(ByVal UserIndex As Integer, ByRef Buffer As Network.Writer, Optional ByVal ValidateInvi As Boolean = False)
#Else
Private Sub SendToUserAliveAreaButindex(ByVal UserIndex As Integer, ByRef Buffer As clsNetWriter, Optional ByVal ValidateInvi As Boolean = False)
#End If
On Error GoTo SendToUserAliveAreaButindex_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
106     If Not MapaValido(Map) Then Exit Sub
108     With UserList(UserIndex)
110         For LoopC = 1 To ConnGroups(Map).CountEntrys
112             tempIndex = ConnGroups(Map).UserEntrys(LoopC)
114             If tempIndex <> UserIndex Then
116                 If CanSendToUser(UserList(UserIndex), UserList(tempIndex), tempIndex, Buffer, ValidateInvi) Then
118                      Call modNetwork.Send(tempIndex, Buffer)
                    End If
                End If
120         Next LoopC
        End With
        Exit Sub

SendToUserAliveAreaButindex_Err:
124     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAliveAreaButindex", Erl)
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToAdminAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToAdminAreaButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToUserAreaButindex_Err
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

122                 If tempIndex <> UserIndex And EsGM(tempIndex) Then

124                     If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                            If CompararPrivilegios(UserList(tempIndex).flags.Privilegios, UserList(UserIndex).flags.Privilegios) >= 0 Then
126                             Call modNetwork.Send(tempIndex, Buffer)
                            End If

                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToAdminAreaButIndex", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserAreaButGMs(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToUserAreaButGMs(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
        
        On Error GoTo SendToUserAreaButindex_Err
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

124                     If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                            If CompararPrivilegios(UserList(UserIndex).flags.Privilegios, UserList(tempIndex).flags.Privilegios) >= 0 Then
126                             Call modNetwork.Send(tempIndex, Buffer)
                            End If
                        End If

                    End If

                End If

            End If

128     Next LoopC

        
        Exit Sub

SendToUserAreaButindex_Err:
130     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserAreaButindex", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToUserGuildArea_Err
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
120                 If UserList(tempIndex).ConnectionDetails.ConnIDValida And UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Then
122                     Call modNetwork.Send(tempIndex, Buffer)
                        

                    End If

                End If

            End If

124     Next LoopC

        
        Exit Sub

SendToUserGuildArea_Err:
126     Call TraceError(Err.Number, Err.Description, "modSendData.SendToUserGuildArea", Erl)

        
End Sub

#If DIRECT_PLAY = 0 Then
Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToNpcArea_Err
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
122                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then

                        If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                            Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                        End If
                        
124                     Call modNetwork.Send(tempIndex, Buffer)
                    End If

                End If

            End If

126     Next LoopC

        
        Exit Sub

SendToNpcArea_Err:
128     Call TraceError(Err.Number, Err.Description, "modSendData.SendToNpcArea", Erl)

        
End Sub
#If DIRECT_PLAY = 0 Then
Private Sub SendToNpcAliveArea(ByVal NpcIndex As Long, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToNpcAliveArea(ByVal NpcIndex As Long, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToNpcArea_Err
        Dim LoopC     As Long
        Dim TempInt   As Integer
        Dim tempIndex As Integer
        Dim map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        
100     If NpcIndex = 0 Then Exit Sub
        
102     map = NpcList(NpcIndex).Pos.map
104     AreaX = NpcList(NpcIndex).AreasInfo.AreaPerteneceX
106     AreaY = NpcList(NpcIndex).AreasInfo.AreaPerteneceY
    
108     If Not MapaValido(map) Then Exit Sub
    
110     For LoopC = 1 To ConnGroups(map).CountEntrys
112         tempIndex = ConnGroups(map).UserEntrys(LoopC)
        
114         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

116         If TempInt Then  'Esta en el area?
118             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

120             If TempInt Then
122                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        If UserList(tempIndex).flags.Muerto = 0 Then
                            If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                                Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                            End If
                            
124                         Call modNetwork.Send(tempIndex, Buffer)
                        End If
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

#If DIRECT_PLAY = 0 Then
        Dim Buffer As Network.Writer
        Set Buffer = Protocol_Writes.GetWriterBuffer()
#Else
        Dim Buffer As clsNetWriter
        Set Buffer = Protocol_Writes.Writer
#End If
        
106     For LoopC = 1 To ConnGroups(Map).CountEntrys
108         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
           
110         TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

112         If TempInt Then  'Esta en el area?
114             TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

116             If TempInt Then
118                 If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
120                     Call modNetwork.Send(tempIndex, Buffer)
                        If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                            Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                        End If

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
#If DIRECT_PLAY = 0 Then
Private Sub SendToMap(ByVal Map As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToMap(ByVal Map As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToMap_Err
        Dim LoopC     As Long
        Dim tempIndex As Integer
100     If Not MapaValido(Map) Then Exit Sub
102     For LoopC = 1 To ConnGroups(Map).CountEntrys
104         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
106         If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
108             Call modNetwork.Send(tempIndex, Buffer)
                If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
                    Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                End If
            End If

110     Next LoopC

        
        Exit Sub

SendToMap_Err:
112     Call TraceError(Err.Number, Err.Description, "modSendData.SendToMap", Erl)

        
End Sub
#If DIRECT_PLAY = 0 Then
Private Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
        
        On Error GoTo SendToMapButIndex_Err
        Dim LoopC     As Long
        Dim Map       As Integer
        Dim tempIndex As Integer
        
100     If UserIndex = 0 Then Exit Sub
        
102     Map = UserList(UserIndex).Pos.Map
    
104     If Not MapaValido(Map) Then Exit Sub

106     For LoopC = 1 To ConnGroups(Map).CountEntrys
108         tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
110         If tempIndex <> UserIndex And UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                If IsValidUserRef(UserList(tempIndex).flags.GMMeSigue) Then
112                 Call modNetwork.Send(UserList(tempIndex).flags.GMMeSigue.ArrayIndex, Buffer)
                End If
113             Call modNetwork.Send(tempIndex, Buffer)

            End If

114     Next LoopC

        
        Exit Sub

SendToMapButIndex_Err:
116     Call TraceError(Err.Number, Err.Description, "modSendData.SendToMapButIndex", Erl)

        
End Sub
#If DIRECT_PLAY = 0 Then
Private Sub SendToGroup(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToGroup(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToGroup_Err
        Dim LoopC     As Long
100     If UserIndex = 0 Then Exit Sub
        If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
        With UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo
106         For LoopC = 1 To .CantidadMiembros
                If IsValidUserRef(.Miembros(LoopC)) Then
                    Call modNetwork.Send(.Miembros(LoopC).ArrayIndex, Buffer)
                End If
            Next LoopC
        End With
        Exit Sub
SendToGroup_Err:
    Call TraceError(Err.Number, Err.Description, "modSendData.SendToGroup", Erl)
End Sub
#If DIRECT_PLAY = 0 Then
Private Sub SendToGroupButIndex(ByVal UserIndex As Integer, ByVal Buffer As Network.Writer)
#Else
Private Sub SendToGroupButIndex(ByVal UserIndex As Integer, ByVal Buffer As clsNetWriter)
#End If
On Error GoTo SendToGroupButIndex_Err
        Dim LoopC     As Long
100     If UserIndex = 0 Then Exit Sub
        If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
        With UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo
106         For LoopC = 1 To .CantidadMiembros
                If IsValidUserRef(.Miembros(LoopC)) And .Miembros(LoopC).ArrayIndex <> UserIndex Then
                    Call modNetwork.Send(.Miembros(LoopC).ArrayIndex, Buffer)
                End If
            Next LoopC
        End With
        Exit Sub
SendToGroupButIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modSendData.SendToGroupButIndex", Erl)
End Sub
