Attribute VB_Name = "modShipTravel"
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
'    Copyright (C) 2002 Mrquez Pablo Ignacio
Option Explicit

Private Const SHIP_TRAVEL_INTERVAL_MS As Long = 12000
Private Const SHIP_TRAVEL_TIME_LIMIT_MS As Long = 100

Private m_LastShipTravelTick As Long

Public Sub ResetShipTravelTimer()
    m_LastShipTravelTick = GetTickCountRaw()
End Sub

Public Sub MaybeRunShipTravel()
    On Error GoTo Handler

    If Not IsFeatureEnabled("ShipTravelEnabled") Then Exit Sub

    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()

    If m_LastShipTravelTick = 0 Then
        m_LastShipTravelTick = nowRaw
        Exit Sub
    End If

    If TicksElapsed(m_LastShipTravelTick, nowRaw) < SHIP_TRAVEL_INTERVAL_MS Then Exit Sub

    m_LastShipTravelTick = nowRaw

    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Call UpdateBarcoForgatNix
    Call UpdateBarcoNixArghal
    Call UpdateBarcoArghalForgat
    Call MsnEnbarque(ForgatDock)
    Call MsnEnbarque(ArghalDock)
    Call MsnEnbarque(NixDock)
    Call PerformTimeLimitCheck(PerformanceTimer, "MaybeRunShipTravel", SHIP_TRAVEL_TIME_LIMIT_MS)
    Exit Sub

Handler:
    Call TraceError(Err.Number, Err.Description, "modShipTravel.MaybeRunShipTravel")
End Sub

Private Function GetPassSlot(ByVal UserIndex As Integer) As Integer
    Dim i As Integer
    With UserList(UserIndex)
        For i = 1 To UBound(.invent.Object)
            ' Le saco el item requerido de Forgat a Nix
            ' Es el mismo item que de Nix a Arghal y de Arghal a Forgat
            If .invent.Object(i).ObjIndex = BarcoNavegandoForgatNix.RequiredPassID Then
                GetPassSlot = i
                Exit Function
            End If
        Next
    End With
    GetPassSlot = -1
End Function

Private Sub MsnEnbarque(ByRef ShipInfo As t_Transport)
    On Error GoTo SendToMap_Err
    Dim LoopC     As Long
    Dim tempIndex As Integer
    If Not MapaValido(ShipInfo.Map) Then Exit Sub
    For LoopC = 1 To ConnGroups(ShipInfo.Map).CountEntrys
        tempIndex = ConnGroups(ShipInfo.Map).UserEntrys(LoopC)
        If UserList(tempIndex).ConnectionDetails.ConnIDValida And UserList(tempIndex).pos.x >= ShipInfo.startX And UserList(tempIndex).pos.x <= ShipInfo.EndX And UserList( _
                tempIndex).pos.y >= ShipInfo.startY And UserList(tempIndex).pos.y <= ShipInfo.EndY Then
            If Not GetPassSlot(tempIndex) > 0 Then
                Call WriteLocaleMsg(tempIndex, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
            Else
                Call WriteLocaleMsg(tempIndex, MsgStartingTrip, e_FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
    Next LoopC
    Exit Sub
SendToMap_Err:
    Call TraceError(Err.Number, Err.Description, "modSendData.SendToMap", Erl)
End Sub

Private Sub UpdateBarcoForgatNix()
    Dim TileX, TileY As Integer
    Dim User     As Integer
    Dim PassSlot As Integer
    ' Modificado por Shugar 5/6/24
    ' Viaje de Forgat a Nix
    ' Verificar si el barco esta en un muelle:
    ' Para ver si esta en el muelle o no, miramos hay un NpcIndex en Map DockX DockY del mapa BarcoNavegando.
    ' Ese Npc solia ser un muelle, y se usa de referencia para saber si el barco partio o sigue quieto.
    ' Si no hay NPC ahi es que el barco esta navegando, por lo tanto no hay movimiento de pasajeros.
    If MapData(BarcoNavegandoForgatNix.Map, BarcoNavegandoForgatNix.DockX, BarcoNavegandoForgatNix.DockY).NpcIndex = 0 Then
        Exit Sub
    End If
    ' Desembarcar: bajamos del barco a los usuarios que llegan a Nix
    ' Para cada tile en el area del barco
    For TileX = BarcoNavegandoForgatNix.startX To BarcoNavegandoForgatNix.EndX
        For TileY = BarcoNavegandoForgatNix.startY To BarcoNavegandoForgatNix.EndY
            ' Si hay un usuario en el tile del barco
            User = MapData(BarcoNavegandoForgatNix.Map, TileX, TileY).UserIndex
            If User > 0 Then
                ' Enviar usuario a Nix
                Call WarpToLegalPos(User, NixDock.Map, NixDock.DestX, NixDock.DestY, True)
                Call WriteLocaleMsg(User, MsgThanksForTravelNix, e_FontTypeNames.FONTTYPE_GUILD)
            End If
        Next TileY
    Next TileX
    ' Embarcacar: subimos al barco a los usuarios que salen de Forgat
    ' Para cada tile en el area del muelle de Forgat
    For TileX = ForgatDock.startX To ForgatDock.EndX
        For TileY = ForgatDock.startY To ForgatDock.EndY
            ' Si hay un usuario en el tile del muelle
            User = MapData(ForgatDock.Map, TileX, TileY).UserIndex
            If User > 0 Then
                ' Sacarle el pasaje y moverlo al barco navegando
                PassSlot = GetPassSlot(User)
                If PassSlot > 0 Then
                    Call WriteLocaleMsg(User, MsgPassForgat, e_FontTypeNames.FONTTYPE_GUILD)
                    Call QuitarUserInvItem(User, PassSlot, 1)
                    Call UpdateUserInv(False, User, PassSlot)
                    Call WarpToLegalPos(User, BarcoNavegandoForgatNix.Map, BarcoNavegandoForgatNix.DestX, BarcoNavegandoForgatNix.DestY, True)
                Else
                    Call WriteLocaleMsg(User, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        Next TileY
    Next TileX
End Sub

Private Sub UpdateBarcoNixArghal()
    Dim TileX, TileY As Integer
    Dim User     As Integer
    Dim PassSlot As Integer
    If NixDock.Map > NumMaps Then Exit Sub
    ' Modificado por Shugar 5/6/24
    ' Viaje de Nix a Arghal
    ' Verificar si el barco esta en un muelle:
    ' Para ver si esta en el muelle o no, miramos hay un NpcIndex en Map DockX DockY del mapa BarcoNavegando.
    ' Ese Npc solia ser un muelle, y se usa de referencia para saber si el barco partio o sigue quieto.
    ' Si no hay NPC ahi es que el barco esta navegando, por lo tanto no hay movimiento de pasajeros.
    If MapData(BarcoNavegandoNixArghal.Map, BarcoNavegandoNixArghal.DockX, BarcoNavegandoNixArghal.DockY).NpcIndex = 0 Then
        Exit Sub
    End If
    ' Desembarcar: bajamos del barco a los usuarios que llegan a Arghal
    ' Para cada tile en el area del barco
    For TileX = BarcoNavegandoNixArghal.startX To BarcoNavegandoNixArghal.EndX
        For TileY = BarcoNavegandoNixArghal.startY To BarcoNavegandoNixArghal.EndY
            ' Si hay un usuario en el tile del barco
            User = MapData(BarcoNavegandoNixArghal.Map, TileX, TileY).UserIndex
            If User > 0 Then
                ' Enviar al usuario a Arghal
                Call WarpToLegalPos(User, ArghalDock.Map, ArghalDock.DestX, ArghalDock.DestY, True)
                Call WriteLocaleMsg(User, MsgThanksForTravelArghal, e_FontTypeNames.FONTTYPE_GUILD)
            End If
        Next TileY
    Next TileX
    ' Embarcacar: subimos al barco a los usuarios que salen de Nix
    ' Para cada tile en el area del muelle de Nix
    For TileX = NixDock.startX To NixDock.EndX
        For TileY = NixDock.startY To NixDock.EndY
            ' Si hay un usuario en el tile del muelle
            User = MapData(NixDock.Map, TileX, TileY).UserIndex
            If User > 0 Then
                ' Sacarle el pasaje y moverlo al barco navegando
                PassSlot = GetPassSlot(User)
                If PassSlot > 0 Then
                    Call WriteLocaleMsg(User, MsgPassNix, e_FontTypeNames.FONTTYPE_GUILD)
                    Call QuitarUserInvItem(User, PassSlot, 1)
                    Call UpdateUserInv(False, User, PassSlot)
                    Call WarpToLegalPos(User, BarcoNavegandoNixArghal.Map, BarcoNavegandoNixArghal.DestX, BarcoNavegandoNixArghal.DestY, True)
                Else
                    Call WriteLocaleMsg(User, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        Next TileY
    Next TileX
End Sub

Private Sub UpdateBarcoArghalForgat()
    Dim TileX, TileY As Integer
    Dim User     As Integer
    Dim PassSlot As Integer
    If ArghalDock.Map > NumMaps Then Exit Sub
    ' Modificado por Shugar 5/6/24
    ' Viaje de Arghal a Forgat
    ' Verificar si el barco esta en un muelle:
    ' Para ver si esta en el muelle o no, miramos hay un NpcIndex en Map DockX DockY del mapa BarcoNavegando.
    ' Ese Npc solia ser un muelle, y se usa de referencia para saber si el barco partio o sigue quieto.
    ' Si no hay NPC ahi es que el barco esta navegando, por lo tanto no hay movimiento de pasajeros.
    If MapData(BarcoNavegandoArghalForgat.Map, BarcoNavegandoArghalForgat.DockX, BarcoNavegandoArghalForgat.DockY).NpcIndex = 0 Then
        Exit Sub
    End If
    ' Desembarcar: bajamos del barco a los usuarios que llegan a Forgat
    ' Para cada tile en el area del barco
    For TileX = BarcoNavegandoArghalForgat.startX To BarcoNavegandoArghalForgat.EndX
        For TileY = BarcoNavegandoArghalForgat.startY To BarcoNavegandoArghalForgat.EndY
            ' Si hay un usuario en el tile del barco
            User = MapData(BarcoNavegandoArghalForgat.Map, TileX, TileY).UserIndex
            If User > 0 Then
                ' Enviar al usuario a Forgat
                Call WarpToLegalPos(User, ForgatDock.Map, ForgatDock.DestX, ForgatDock.DestY, True)
                Call WriteLocaleMsg(User, MsgThanksForTravelForgat, e_FontTypeNames.FONTTYPE_GUILD)
            End If
        Next TileY
    Next TileX
    ' Embarcacar: subimos al barco a los usuarios que salen de Arghal
    ' Para cada tile en el area del muelle de Arghal
    For TileX = ArghalDock.startX To ArghalDock.EndX
        For TileY = ArghalDock.startY To ArghalDock.EndY
            ' Si hay un usuario en el tile del muelle
            User = MapData(ArghalDock.Map, TileX, TileY).UserIndex
            If User > 0 Then
                ' Sacarle el pasaje y moverlo al barco navegando
                PassSlot = GetPassSlot(User)
                If PassSlot > 0 Then
                    Call WriteLocaleMsg(User, MsgPassArghal, e_FontTypeNames.FONTTYPE_GUILD)
                    Call QuitarUserInvItem(User, PassSlot, 1)
                    Call UpdateUserInv(False, User, PassSlot)
                    Call WarpToLegalPos(User, BarcoNavegandoArghalForgat.Map, BarcoNavegandoArghalForgat.DestX, BarcoNavegandoArghalForgat.DestY, True)
                Else
                    Call WriteLocaleMsg(User, MsgInvalidPass, e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If
        Next TileY
    Next TileX
End Sub
