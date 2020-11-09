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

'Option Explicit

Public Enum SendTarget

    ToAll = 1
    toMap
    ToPCArea
    ToAllButIndex
    ToMapButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
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
    
    Select Case sndRoute

        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
            
        Case SendTarget.ToSuperiores

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
                        If UserList(LoopC).flags.Privilegios >= PlayerType.Admin Then
                            Call EnviarDatosASlot(LoopC, sndData)

                        End If

                    End If

                End If

            Next LoopC

            Exit Sub
            
        Case SendTarget.ToAll

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToAllButIndex

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
          
        Case SendTarget.ToMapButIndex
            Call SendToMapButIndex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)

                End If

                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call EnviarDatosASlot(LoopC, sndData)

                End If

                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
        Case SendTarget.ToConsejo

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToConsejoCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToRolesMasters

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCiudadanos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If Status(LoopC) < 2 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCriminales

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If Status(LoopC) = 2 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToReal

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCiudadanosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If Status(LoopC) < 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCriminalesYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If Status(LoopC) = 2 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToRealYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCaosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call EnviarDatosASlot(LoopC, sndData)

                    End If

                End If

            Next LoopC

            Exit Sub

    End Select

End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

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
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)

        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next LoopC

End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)

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
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    'sdData = sdData & ENDC

    If Not MapaValido(Map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
            
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If tempIndex <> UserIndex Then
                    If UserList(tempIndex).ConnIDValida Then
                        Call EnviarDatosASlot(tempIndex, sdData)

                    End If

                End If

            End If

        End If

    Next LoopC

End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

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
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then

                'Dead and admins read
                If UserList(tempIndex).ConnIDValida = True And (UserList(tempIndex).flags.Muerto = 1 Or (UserList(tempIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next LoopC

End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)

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
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    If UserList(UserIndex).GuildIndex = 0 Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida And UserList(tempIndex).GuildIndex = UserList(UserIndex).GuildIndex Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next LoopC

End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, ByVal sdData As String)

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
    
    Map = UserList(UserIndex).Pos.Map
    AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
    AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
    
    If Not MapaValido(Map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
            If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                If UserList(tempIndex).ConnIDValida Then
                    If UserList(tempIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next LoopC

End Sub

Private Sub SendToNpcArea(ByVal NpcIndex As Long, ByVal sdData As String)

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
    
    Map = Npclist(NpcIndex).Pos.Map
    AreaX = Npclist(NpcIndex).AreasInfo.AreaPerteneceX
    AreaY = Npclist(NpcIndex).AreasInfo.AreaPerteneceY
    'sdData = sdData & ENDC
    
    If Not MapaValido(Map) Then Exit Sub
    
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next LoopC

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal AreaX As Integer, ByVal AreaY As Integer, ByVal sdData As String)
 
    Dim LoopC     As Long

    Dim TempInt   As Integer

    Dim tempIndex As Integer
   
    AreaX = 2 ^ (AreaX \ 12)
    AreaY = 2 ^ (AreaY \ 12)
   
    If Not MapaValido(Map) Then Exit Sub
 
    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
           
        TempInt = UserList(tempIndex).AreasInfo.AreaReciveX And AreaX

        If TempInt Then  'Esta en el area?
            TempInt = UserList(tempIndex).AreasInfo.AreaReciveY And AreaY

            If TempInt Then
                If UserList(tempIndex).ConnIDValida Then
                    Call EnviarDatosASlot(tempIndex, sdData)

                End If

            End If

        End If

    Next LoopC

End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim LoopC     As Long

    Dim tempIndex As Integer
    
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)

        End If

    Next LoopC

End Sub

Public Sub SendToMapButIndex(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim LoopC     As Long

    Dim Map       As Integer

    Dim tempIndex As Integer
    
    Map = UserList(UserIndex).Pos.Map
    
    If Not MapaValido(Map) Then Exit Sub

    For LoopC = 1 To ConnGroups(Map).CountEntrys
        tempIndex = ConnGroups(Map).UserEntrys(LoopC)
        
        If tempIndex <> UserIndex And UserList(tempIndex).ConnIDValida Then
            Call EnviarDatosASlot(tempIndex, sdData)

        End If

    Next LoopC

End Sub

