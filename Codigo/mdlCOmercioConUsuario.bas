Attribute VB_Name = "mdlCOmercioConUsuario"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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



'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Function IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer) As Boolean
    On Error GoTo ErrHandler
    'Si ambos pusieron /comerciar entonces
    If UserList(Origen).ComUsu.DestUsu.ArrayIndex = Destino And UserList(Destino).ComUsu.DestUsu.ArrayIndex = Origen Then
        If UserList(Origen).pos.Map <> UserList(Destino).pos.Map Then
            Call WriteLocaleMsg(Origen, MSG_NO_COMERCIO_CANCEL_PORQUE_ESTN_MISMO_MAPA, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja) 'Msg2108= El comercio se cancel porque ya no estn en el mismo mapa.
            Call WriteLocaleMsg(Destino, MSG_NO_COMERCIO_CANCEL_PORQUE_ESTN_MISMO_MAPA, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja) 'Msg2108= El comercio se cancel porque ya no estn en el mismo mapa.
            Call FinComerciarUsu(Origen, True)
            Call FinComerciarUsu(Destino, True)
            IniciarComercioConUsuario = False
            Exit Function
        End If
        'Actualiza el inventario del usuario
        Call UpdateUserInv(True, Origen, 0)
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Origen)
        UserList(Origen).flags.Comerciando = True
        'Actualiza el inventario del usuario
        Call UpdateUserInv(True, Destino, 0)
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Destino)
        UserList(Destino).flags.Comerciando = True
        'Limpio los arrays antes de iniciar el comercio seguro.
        Erase UserList(Origen).ComUsu.itemsAenviar
        Erase UserList(Destino).ComUsu.itemsAenviar
        UserList(Destino).ComUsu.Oro = 0
        UserList(Origen).ComUsu.Oro = 0
        UserList(Origen).ComUsu.Acepto = False
        UserList(Destino).ComUsu.Acepto = False
        Call ClearUserRef(UserList(Origen).ComUsu.InvitationFrom)
        Call ClearUserRef(UserList(Destino).ComUsu.InvitationFrom)
        If UserList(Origen).flags.pregunta = 4 Then
            UserList(Origen).flags.pregunta = 0
            UserList(Origen).flags.RespondiendoPregunta = False
        End If
        If UserList(Destino).flags.pregunta = 4 Then
            UserList(Destino).flags.pregunta = 0
            UserList(Destino).flags.RespondiendoPregunta = False
        End If
        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        'Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", e_FontTypeNames.FONTTYPE_TALK)
        Call SetUserRef(UserList(Destino).flags.TargetUser, Origen)
        Call SetUserRef(UserList(Destino).ComUsu.InvitationFrom, Origen)
        UserList(Destino).flags.pregunta = 4
        UserList(Destino).flags.RespondiendoPregunta = True
        Call WritePreguntaBox(Destino, MSG_DESEA_COMERCIAR_CONTIGO_ACEPTAS, UserList(Origen).name) 'Msg1594= ¬1 desea comerciar contigo. ¿Aceptás?
    End If
    IniciarComercioConUsuario = True
    Exit Function
ErrHandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)
End Function

Public Sub AcceptSafeTradeInvitation(ByVal UserIndex As Integer)
    On Error GoTo AcceptSafeTradeInvitation_Err
    Dim invitation As t_UserReference
    Dim TargetIndex As Integer
    With UserList(UserIndex)
        If .flags.pregunta <> 4 Then Exit Sub
        .flags.RespondiendoPregunta = False
        invitation = .ComUsu.InvitationFrom
        Call ClearUserRef(.ComUsu.InvitationFrom)
        .flags.pregunta = 0
        If IsValidUserRef(invitation) Then
            TargetIndex = invitation.ArrayIndex
            If SafeTradeInvitationMatches(UserList(TargetIndex).ComUsu, UserIndex, .VersionId) And _
                    UserList(TargetIndex).flags.UserLogged And Not UserList(TargetIndex).flags.Comerciando And _
                    Not .flags.Comerciando And .flags.Muerto = 0 And UserList(TargetIndex).flags.Muerto = 0 And _
                    .pos.Map = UserList(TargetIndex).pos.Map And Distancia(.pos, UserList(TargetIndex).pos) <= 3 And MapInfo(.pos.Map).Seguro <> 0 Then
                .ComUsu.DestUsu = invitation
                .ComUsu.DestNick = GetUserRealName(TargetIndex)
                .ComUsu.cant = 0
                .ComUsu.Objeto = 0
                .ComUsu.Acepto = False
                Call IniciarComercioConUsuario(UserIndex, TargetIndex)
            Else
                If SafeTradeInvitationMatches(UserList(TargetIndex).ComUsu, UserIndex, .VersionId) And Not UserList(TargetIndex).flags.Comerciando Then
                    Call ClearSafeTradeRequest(UserList(TargetIndex).ComUsu)
                End If
                Call WriteLocaleMsg(UserIndex, MSG_SERVIDOR_SOLICITUD_COMERCIO_INVALIDA_REINTENTE, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
            End If
        Else
            Call WriteLocaleMsg(UserIndex, MSG_SERVIDOR_SOLICITUD_COMERCIO_INVALIDA_REINTENTE, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER)
        End If
    End With
    Exit Sub
AcceptSafeTradeInvitation_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.AcceptSafeTradeInvitation", Erl)
End Sub

Public Sub RejectSafeTradeInvitation(ByVal UserIndex As Integer)
    On Error GoTo RejectSafeTradeInvitation_Err
    Dim invitation As t_UserReference
    Dim TargetIndex As Integer
    With UserList(UserIndex)
        If .flags.pregunta <> 4 Then Exit Sub
        .flags.RespondiendoPregunta = False
        invitation = .ComUsu.InvitationFrom
        Call ClearUserRef(.ComUsu.InvitationFrom)
        .flags.pregunta = 0
        If IsValidUserRef(invitation) Then
            TargetIndex = invitation.ArrayIndex
            If SafeTradeInvitationMatches(UserList(TargetIndex).ComUsu, UserIndex, .VersionId) And Not UserList(TargetIndex).flags.Comerciando Then
                Call WriteLocaleMsg(TargetIndex, MSG_EL_USUARIO_NO_DESEA_COMERCIAR_EN_ESTE_MOMENTO, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_INFO)
                Call ClearSafeTradeRequest(UserList(TargetIndex).ComUsu)
            End If
        End If
    End With
    Exit Sub
RejectSafeTradeInvitation_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.RejectSafeTradeInvitation", Erl)
End Sub

Public Function SafeTradeInvitationMatches(ByRef request As t_ComercioUsuario, ByVal recipient As Integer, ByVal recipientVersion As Integer) As Boolean
    On Error GoTo SafeTradeInvitationMatches_Err
    SafeTradeInvitationMatches = recipient > 0 And request.DestUsu.ArrayIndex = recipient And request.DestUsu.VersionId = recipientVersion
    Exit Function
SafeTradeInvitationMatches_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.SafeTradeInvitationMatches", Erl)
End Function

Public Sub ClearSafeTradeRequest(ByRef request As t_ComercioUsuario)
    On Error GoTo ClearSafeTradeRequest_Err
    Call ClearUserRef(request.DestUsu)
    Call ClearUserRef(request.InvitationFrom)
    request.DestNick = vbNullString
    request.Objeto = 0
    request.cant = 0
    request.Oro = 0
    request.Acepto = False
    Erase request.itemsAenviar
    Exit Sub
ClearSafeTradeRequest_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.ClearSafeTradeRequest", Erl)
End Sub

Public Function SafeTradeOfferedAmount(ByRef items() As t_Obj, ByVal objectIndex As Integer, ByVal tags As Long) As Long
    On Error GoTo SafeTradeOfferedAmount_Err
    Dim i As Long
    For i = LBound(items) To UBound(items)
        If items(i).ObjIndex = objectIndex And items(i).ElementalTags = tags Then
            SafeTradeOfferedAmount = SafeTradeOfferedAmount + items(i).amount
        End If
    Next i
    Exit Function
SafeTradeOfferedAmount_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.SafeTradeOfferedAmount", Erl)
End Function

' Apply an addition atomically. A full board cannot silently drop a remainder
' or lose elemental tags when splitting a stack between two offer positions.
Public Function AddSafeTradeOffer(ByRef items() As t_Obj, ByRef offeredGold As Long, ByRef item As t_Obj, ByVal availableGold As Long, ByVal maxStack As Long) As Boolean
    On Error GoTo AddSafeTradeOffer_Err
    If item.amount <= 0 Or maxStack <= 0 Then Exit Function
    If item.ObjIndex = 0 Then
        If item.amount > availableGold - offeredGold Then Exit Function
        offeredGold = offeredGold + item.amount
        AddSafeTradeOffer = True
        Exit Function
    End If
    Dim proposed() As t_Obj
    Dim i          As Long
    Dim remaining  As Long
    Dim addition   As Long
    ReDim proposed(LBound(items) To UBound(items))
    For i = LBound(items) To UBound(items)
        proposed(i) = items(i)
    Next i
    remaining = item.amount
    For i = LBound(proposed) To UBound(proposed)
        If proposed(i).ObjIndex = item.ObjIndex And proposed(i).ElementalTags = item.ElementalTags Then
            addition = min(remaining, maxStack - proposed(i).amount)
            If addition > 0 Then
                proposed(i).amount = proposed(i).amount + addition
                remaining = remaining - addition
            End If
        End If
    Next i
    For i = LBound(proposed) To UBound(proposed)
        If remaining > 0 And proposed(i).ObjIndex = 0 Then
            proposed(i) = item
            proposed(i).amount = min(remaining, maxStack)
            remaining = remaining - proposed(i).amount
        End If
    Next i
    If remaining > 0 Then Exit Function
    For i = LBound(items) To UBound(items)
        items(i) = proposed(i)
    Next i
    AddSafeTradeOffer = True
    Exit Function
AddSafeTradeOffer_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.AddSafeTradeOffer", Erl)
End Function

Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer, ByVal UserIndex As Integer, ByRef ObjAEnviar As t_Obj)
    On Error GoTo EnviarObjetoTransaccion_Err
    Dim hasItems As Boolean
    hasItems = True
    If ObjAEnviar.ObjIndex > 0 Then
        hasItems = TieneObjetos(ObjAEnviar.ObjIndex, SafeTradeOfferedAmount(UserList(UserIndex).ComUsu.itemsAenviar, ObjAEnviar.ObjIndex, ObjAEnviar.ElementalTags) + ObjAEnviar.amount, UserIndex, ObjAEnviar.ElementalTags)
    End If
    If Not hasItems Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ESA_CANTIDAD_DISPONIBLE_AGREGAR_1997, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja, vbNullString)
    ElseIf Not AddSafeTradeOffer(UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, ObjAEnviar, UserList(UserIndex).Stats.GLD, GetMaxInvOBJ()) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_SUFICIENTE_LUGAR_AGREGAR_ESA_CANTIDAD, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja, vbNullString)
    Else
        ' Keep the existing accept packet and server authority. An addition
        ' invalidates the OTHER player's prior acceptance before publishing it.
        If UserList(AQuien).ComUsu.Acepto Then
            UserList(AQuien).ComUsu.Acepto = False
            Call WriteLocaleMsg(AQuien, MSG_CAMBIADO_OFERTA, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_PROMEDIO_MAYOR, GetUserDisplayName(UserIndex))
        End If
        Call WriteChangeUserTradeSlot(AQuien, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, False)
    End If
    ' Echo even an unchanged board after a rejected addition.
    Call WriteChangeUserTradeSlot(UserIndex, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, True)
    Exit Sub
EnviarObjetoTransaccion_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.EnviarObjetoTransaccion", Erl)
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer, Optional ByVal Invalido As Boolean = False)
    On Error GoTo FinComerciarUsu_Err
    If UserIndex = 0 Then Exit Sub
    With UserList(UserIndex)
        If IsValidUserRef(.ComUsu.DestUsu) And Not Invalido Then
            Call WriteUserCommerceEnd(UserIndex)
        End If
        Call ClearSafeTradeRequest(.ComUsu)
        If .flags.pregunta = 4 Then
            .flags.pregunta = 0
            .flags.RespondiendoPregunta = False
        End If
        .flags.Comerciando = False
    End With
    Exit Sub
FinComerciarUsu_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.FinComerciarUsu", Erl)
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
    On Error GoTo AceptarComercioUsu_Err
    Dim objOfrecido   As t_Obj
    Dim OtroUserIndex As Integer
    Dim TerminarAhora As Boolean
    TerminarAhora = UserList(UserIndex).ComUsu.DestUsu.ArrayIndex <= 0 Or UserList(UserIndex).ComUsu.DestUsu.ArrayIndex > MaxUsers
    OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu.ArrayIndex
    If Not TerminarAhora Then
        TerminarAhora = Not UserList(OtroUserIndex).flags.UserLogged Or Not UserList(UserIndex).flags.UserLogged
    End If
    If Not TerminarAhora Then
        TerminarAhora = UserList(OtroUserIndex).ComUsu.DestUsu.ArrayIndex <> UserIndex
    End If
    If TerminarAhora Then
        Call FinComerciarUsu(UserIndex)
        If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
            Call FinComerciarUsu(OtroUserIndex)
        End If
        Exit Sub
    End If
    UserList(UserIndex).ComUsu.Acepto = True
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        'Call WriteConsoleMsg(UserIndex, "El otro usuario aun no ha aceptado tu oferta.", e_FontTypeNames.FONTTYPE_TALK)
        Call WriteLocaleMsg(UserIndex, MSG_NO_OTRO_USUARIO_AUN_HA_ACEPTADO_OFERTA, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja) 'Msg1596= El otro usuario aún no ha aceptado tu oferta.
        Exit Sub
    End If
    If UserList(UserIndex).ComUsu.Oro > UserList(UserIndex).Stats.GLD Then
        'Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", e_FontTypeNames.FONTTYPE_TALK)'ver ReyarB
        Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ESA_CANTIDAD_1597, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja) 'Msg1597= No tienes esa cantidad.
        TerminarAhora = True
    End If
    If UserList(OtroUserIndex).ComUsu.Oro > UserList(OtroUserIndex).Stats.GLD Then
        Call WriteLocaleMsg(OtroUserIndex, MSG_NO_TIENES_ESA_CANTIDAD, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja, vbNullString) ' Msg1999=No tienes esa cantidad.
        GoTo FinalizarComercio
    End If
    ' Verificamos que si tiene los objetos JUSTO ANTES de intercambiarlos
    Dim i As Long
    For i = 1 To UBound(UserList(OtroUserIndex).ComUsu.itemsAenviar)
        objOfrecido = UserList(OtroUserIndex).ComUsu.itemsAenviar(i)
        If objOfrecido.ObjIndex > 0 And Not TieneObjetos(objOfrecido.ObjIndex, SafeTradeOfferedAmount(UserList(OtroUserIndex).ComUsu.itemsAenviar, objOfrecido.ObjIndex, objOfrecido.elementalTags), OtroUserIndex, objOfrecido.elementalTags) Then
            Call WriteLocaleMsg(OtroUserIndex, MSG_NO_OTRO_USUARIO_TIENE_ESA_CANTIDAD_DISPONIBLE_OFRECER, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja) 'Msg1599= El otro usuario no tiene esa cantidad disponible para ofrecer.
            GoTo FinalizarComercio
        End If
        objOfrecido = UserList(UserIndex).ComUsu.itemsAenviar(i)
        If objOfrecido.ObjIndex > 0 And Not TieneObjetos(objOfrecido.ObjIndex, SafeTradeOfferedAmount(UserList(UserIndex).ComUsu.itemsAenviar, objOfrecido.ObjIndex, objOfrecido.elementalTags), UserIndex, objOfrecido.elementalTags) Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ESA_CANTIDAD_DISPONIBLE_OFRECER, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja) 'Msg1598= No tienes esa cantidad disponible para ofrecer.
            GoTo FinalizarComercio
        End If
    Next i
    'Por si las moscas...
    If TerminarAhora Then GoTo FinalizarComercio
    'pone el oro directamente en la billetera
    If UserList(OtroUserIndex).ComUsu.Oro > 0 Then
        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Oro
        Call WriteUpdateUserStats(OtroUserIndex)
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Oro
        Call WriteUpdateUserStats(UserIndex)
    End If
    If UserList(UserIndex).ComUsu.Oro > 0 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Oro
        Call WriteUpdateUserStats(UserIndex)
        UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Oro
        Call WriteUpdateUserStats(OtroUserIndex)
    End If
    ' Confirmamos que SI tienen los objetos a comerciar, procedemos con el cambio.
    For i = 1 To UBound(UserList(OtroUserIndex).ComUsu.itemsAenviar)
        objOfrecido = UserList(OtroUserIndex).ComUsu.itemsAenviar(i)
        If objOfrecido.ObjIndex > 0 Then
            If Not MeterItemEnInventario(UserIndex, objOfrecido) Then
                Call TirarItemAlPiso(UserList(UserIndex).pos, objOfrecido)
            End If
            If QuitarObjetos(objOfrecido.ObjIndex, objOfrecido.amount, OtroUserIndex, objOfrecido.elementalTags) Then
                Call LogSafeCommerceTransfer(GetUserRealName(OtroUserIndex), GetUserRealName(UserIndex), objOfrecido.ObjIndex, objOfrecido.amount, objOfrecido.elementalTags)
            End If
        End If
    Next i
    Dim j As Long
    For j = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
        objOfrecido = UserList(UserIndex).ComUsu.itemsAenviar(j)
        If objOfrecido.ObjIndex > 0 Then
            If MeterItemEnInventario(OtroUserIndex, objOfrecido) = False Then
                Call TirarItemAlPiso(UserList(OtroUserIndex).pos, objOfrecido)
            End If
            If QuitarObjetos(objOfrecido.ObjIndex, objOfrecido.amount, UserIndex, objOfrecido.elementalTags) Then
                Call LogSafeCommerceTransfer(GetUserRealName(UserIndex), GetUserRealName(OtroUserIndex), objOfrecido.ObjIndex, objOfrecido.amount, objOfrecido.elementalTags)
            End If
        End If
    Next j
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserInv(True, OtroUserIndex, 0)
    'Msg2290=Comercio finalizado con éxito.
    Call WriteLocaleMsg(UserIndex, MSG_COMERCIO_FINALIZADO_EXITO, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
    Call WriteLocaleMsg(OtroUserIndex, MSG_COMERCIO_FINALIZADO_EXITO, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
FinalizarComercio:
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
AceptarComercioUsu_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.AceptarComercioUsu", Erl)
End Sub
