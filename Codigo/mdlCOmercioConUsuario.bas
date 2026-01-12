Attribute VB_Name = "mdlCOmercioConUsuario"
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
Private Const MAX_ORO_LOGUEABLE As Long = 50000
Private Const MAX_OBJ_LOGUEABLE As Long = 1000

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Function IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer) As Boolean
    On Error GoTo ErrHandler
    'Si ambos pusieron /comerciar entonces
    If UserList(Origen).ComUsu.DestUsu.ArrayIndex = Destino And UserList(Destino).ComUsu.DestUsu.ArrayIndex = Origen Then
        If UserList(Origen).pos.Map <> UserList(Destino).pos.Map Then
            Call WriteLocaleMsg(Origen, 2108, e_FontTypeNames.FONTTYPE_INFO) 'Msg2108= El comercio se cancel porque ya no estn en el mismo mapa.
            Call WriteLocaleMsg(Destino, 2108, e_FontTypeNames.FONTTYPE_INFO) 'Msg2108= El comercio se cancel porque ya no estn en el mismo mapa.
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
        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        'Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", e_FontTypeNames.FONTTYPE_TALK)
        Call SetUserRef(UserList(Destino).flags.TargetUser, Origen)
        UserList(Destino).flags.pregunta = 4
        Call WritePreguntaBox(Destino, 1594, UserList(Origen).name) 'Msg1594= ¬1 desea comerciar contigo. ¿Aceptás?
    End If
    IniciarComercioConUsuario = True
    Exit Function
ErrHandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)
End Function

Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer, ByVal UserIndex As Integer, ByRef ObjAEnviar As t_Obj)
    On Error GoTo EnviarObjetoTransaccion_Err
    Dim FirstEmptyPos     As Byte
    Dim FoundPos          As Byte
    Dim nada              As Boolean
    Dim cantidadTotalItem As Long
    'Me fijo si recibe oro
    If ObjAEnviar.ObjIndex = 0 Then
        'Si es oro simplemente me fijo si ya había agregado antes y se lo sumo
        If UserList(UserIndex).ComUsu.Oro + ObjAEnviar.amount <= UserList(UserIndex).Stats.GLD Then
            UserList(UserIndex).ComUsu.Oro = UserList(UserIndex).ComUsu.Oro + ObjAEnviar.amount
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1936, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1936=No tienes esa cantidad disponible para agregar.
            Exit Sub
        End If
    Else
        Dim j As Long
        'me fijo si tiene esas cantidades para que no duplique items
        For j = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
            If UserList(UserIndex).ComUsu.itemsAenviar(j).ObjIndex = ObjAEnviar.ObjIndex And UserList(UserIndex).ComUsu.itemsAenviar(j).ElementalTags = ObjAEnviar.ElementalTags _
                    Then
                cantidadTotalItem = cantidadTotalItem + UserList(UserIndex).ComUsu.itemsAenviar(j).amount
            End If
        Next j
        cantidadTotalItem = cantidadTotalItem + ObjAEnviar.amount
        If Not TieneObjetos(ObjAEnviar.ObjIndex, cantidadTotalItem, UserIndex, ObjAEnviar.ElementalTags) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1997, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1997=No tienes esa cantidad disponible para agregar.
            Exit Sub
        End If
        'Si es un item recorro todo el array para ver si ese elemento ya está agregado y de paso me guardo la primer posición vacía
        Dim i As Long
        For i = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
            'Si encuentro el item y tiene lugar pongo Found en la posición que lo encontré
            If UserList(UserIndex).ComUsu.itemsAenviar(i).ObjIndex = ObjAEnviar.ObjIndex And UserList(UserIndex).ComUsu.itemsAenviar(i).ElementalTags = ObjAEnviar.ElementalTags _
                    And UserList(UserIndex).ComUsu.itemsAenviar(i).amount <= 10000 Then
                'Me fijo si le va a entrar el objeto con las cantidades en el slot que encontró
                If UserList(UserIndex).ComUsu.itemsAenviar(i).amount + ObjAEnviar.amount <= GetMaxInvOBJ() Then
                    'Si le entra simplemente le agrego las cantidades
                    UserList(UserIndex).ComUsu.itemsAenviar(i).amount = UserList(UserIndex).ComUsu.itemsAenviar(i).amount + ObjAEnviar.amount
                    nada = True
                    Exit For
                    'Si no le entra la cantidad en ese slot me guardo la posición y mas adelante me fijo si hay otra posición libre.
                Else
                    FoundPos = i
                End If
                'Si no encuentra item en la pos y todavía no guardó ninguna primera posición me la guardo.
            ElseIf UserList(UserIndex).ComUsu.itemsAenviar(i).ObjIndex = 0 And FirstEmptyPos = 0 Then
                FirstEmptyPos = i
            End If
        Next i
        With UserList(UserIndex).ComUsu
            'Si tengo una posición encontrada con un item y a su ves 1 slot vacío para agregar los restantes de ese item
            If FoundPos > 0 And FirstEmptyPos > 0 Then
                Dim restante As Long
                restante = .itemsAenviar(FoundPos).amount + ObjAEnviar.amount - 10000
                If FoundPos > FirstEmptyPos Then
                    .itemsAenviar(FoundPos).amount = restante
                    .itemsAenviar(FirstEmptyPos).amount = 10000
                Else
                    .itemsAenviar(FoundPos).amount = 10000
                    .itemsAenviar(FirstEmptyPos).amount = restante
                End If
                .itemsAenviar(FirstEmptyPos).ObjIndex = ObjAEnviar.ObjIndex
            ElseIf FoundPos = 0 And FirstEmptyPos <> 0 Then
                'Si entré aca es porque tengo que guardar el item en la pos vacía que encontré
                .itemsAenviar(FirstEmptyPos).ObjIndex = ObjAEnviar.ObjIndex
                .itemsAenviar(FirstEmptyPos).amount = ObjAEnviar.amount
                .itemsAenviar(FirstEmptyPos).ElementalTags = ObjAEnviar.ElementalTags
            ElseIf FirstEmptyPos = 0 And nada = False Then
                'le aviso que no le entran los items
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1998, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1998=No tienes suficiente lugar para agregar esa cantidad o item.
            End If
        End With
    End If
    'Le envío la data al cliente para agregar en la lista.
    Call WriteChangeUserTradeSlot(AQuien, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, False)
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
        .ComUsu.Acepto = False
        .ComUsu.cant = 0
        Call SetUserRef(.ComUsu.DestUsu, 0)
        .ComUsu.Objeto = 0
        .ComUsu.DestNick = vbNullString
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
        Call WriteLocaleMsg(UserIndex, 1596, e_FontTypeNames.FONTTYPE_TALK) 'Msg1596= El otro usuario aún no ha aceptado tu oferta.
        Exit Sub
    End If
    If UserList(UserIndex).ComUsu.Oro > UserList(UserIndex).Stats.GLD Then
        'Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", e_FontTypeNames.FONTTYPE_TALK)'ver ReyarB
        Call WriteLocaleMsg(UserIndex, 1597, e_FontTypeNames.FONTTYPE_TALK) 'Msg1597= No tienes esa cantidad.
        TerminarAhora = True
    End If
    If UserList(OtroUserIndex).ComUsu.Oro > UserList(OtroUserIndex).Stats.GLD Then
        Call WriteConsoleMsg(OtroUserIndex, PrepareMessageLocaleMsg(1999, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1999=No tienes esa cantidad.
        GoTo FinalizarComercio
    End If
    ' Verificamos que si tiene los objetos JUSTO ANTES de intercambiarlos
    Dim i As Long
    For i = 1 To UBound(UserList(OtroUserIndex).ComUsu.itemsAenviar)
        objOfrecido = UserList(OtroUserIndex).ComUsu.itemsAenviar(i)
        If objOfrecido.ObjIndex > 0 And Not TieneObjetos(objOfrecido.ObjIndex, objOfrecido.amount, OtroUserIndex, objOfrecido.ElementalTags) Then
            Call WriteLocaleMsg(OtroUserIndex, 1599, e_FontTypeNames.FONTTYPE_INFO) 'Msg1599= El otro usuario no tiene esa cantidad disponible para ofrecer.
            GoTo FinalizarComercio
        End If
        objOfrecido = UserList(UserIndex).ComUsu.itemsAenviar(i)
        If objOfrecido.ObjIndex > 0 And Not TieneObjetos(objOfrecido.ObjIndex, objOfrecido.amount, UserIndex, objOfrecido.ElementalTags) Then
            Call WriteLocaleMsg(UserIndex, 1598, e_FontTypeNames.FONTTYPE_INFO) 'Msg1598= No tienes esa cantidad disponible para ofrecer.
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
        If Not MeterItemEnInventario(UserIndex, UserList(OtroUserIndex).ComUsu.itemsAenviar(i)) Then
            Call TirarItemAlPiso(UserList(UserIndex).pos, UserList(OtroUserIndex).ComUsu.itemsAenviar(i))
        End If
        Call QuitarObjetos(UserList(OtroUserIndex).ComUsu.itemsAenviar(i).ObjIndex, UserList(OtroUserIndex).ComUsu.itemsAenviar(i).amount, OtroUserIndex, UserList( _
                OtroUserIndex).ComUsu.itemsAenviar(i).ElementalTags)
    Next i
    Dim j As Long
    For j = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
        If MeterItemEnInventario(OtroUserIndex, UserList(UserIndex).ComUsu.itemsAenviar(j)) = False Then
            Call TirarItemAlPiso(UserList(OtroUserIndex).pos, UserList(UserIndex).ComUsu.itemsAenviar(j))
        End If
        Call QuitarObjetos(UserList(UserIndex).ComUsu.itemsAenviar(j).ObjIndex, UserList(UserIndex).ComUsu.itemsAenviar(j).amount, UserIndex, UserList( _
                UserIndex).ComUsu.itemsAenviar(j).ElementalTags)
    Next j
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserInv(True, OtroUserIndex, 0)
FinalizarComercio:
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Exit Sub
AceptarComercioUsu_Err:
    Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.AceptarComercioUsu", Erl)
End Sub
