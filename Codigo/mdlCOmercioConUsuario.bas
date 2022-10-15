Attribute VB_Name = "mdlCOmercioConUsuario"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
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

'[Alejo]
Option Explicit

Private Const MAX_ORO_LOGUEABLE As Long = 50000

Private Const MAX_OBJ_LOGUEABLE As Long = 1000


'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Function IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer) As Boolean

        On Error GoTo ErrHandler

        'Si ambos pusieron /comerciar entonces
106     If UserList(Origen).ComUsu.DestUsu.ArrayIndex = Destino And UserList(Destino).ComUsu.DestUsu.ArrayIndex = Origen Then

            'Actualiza el inventario del usuario
108         Call UpdateUserInv(True, Origen, 0)
            'Decirle al origen que abra la ventanita.
110         Call WriteUserCommerceInit(Origen)
112         UserList(Origen).flags.Comerciando = True

            'Actualiza el inventario del usuario
114         Call UpdateUserInv(True, Destino, 0)
            'Decirle al origen que abra la ventanita.
116         Call WriteUserCommerceInit(Destino)
118         UserList(Destino).flags.Comerciando = True
            'Limpio los arrays antes de iniciar el comercio seguro.
120         Erase UserList(Origen).ComUsu.itemsAenviar
122         Erase UserList(Destino).ComUsu.itemsAenviar
124         UserList(Destino).ComUsu.Oro = 0
126         UserList(Origen).ComUsu.Oro = 0
            
            'Call EnviarObjetoTransaccion(Origen)
        Else
            'Es el primero que comercia ?
            'Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", e_FontTypeNames.FONTTYPE_TALK)
128         Call SetUserRef(UserList(Destino).flags.targetUser, Origen)
    
130         UserList(Destino).flags.pregunta = 4
132         Call WritePreguntaBox(Destino, UserList(Origen).name & " desea comerciar contigo. ¿Aceptás?")
    
        End If

        IniciarComercioConUsuario = True

        Exit Function
ErrHandler:
134     Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)

End Function
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer, ByVal UserIndex As Integer, ByRef ObjAEnviar As t_Obj)
        
            On Error GoTo EnviarObjetoTransaccion_Err
        
            Dim FirstEmptyPos As Byte
            Dim FoundPos As Byte
            Dim nada As Boolean
            Dim cantidadTotalItem As Long
        
            'Me fijo si recibe oro
100         If ObjAEnviar.ObjIndex = 0 Then
                'Si es oro simplemente me fijo si ya había agregado antes y se lo sumo
102             If UserList(UserIndex).ComUsu.Oro + ObjAEnviar.amount <= UserList(UserIndex).Stats.GLD Then
104                 UserList(UserIndex).ComUsu.Oro = UserList(UserIndex).ComUsu.Oro + ObjAEnviar.amount
                Else
106                 Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad disponible para agregar.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
        
                Dim j As Long
                'me fijo si tiene esas cantidades para que no duplique items
108             For j = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
110                 If UserList(UserIndex).ComUsu.itemsAenviar(j).ObjIndex = ObjAEnviar.ObjIndex Then
112                     cantidadTotalItem = cantidadTotalItem + UserList(UserIndex).ComUsu.itemsAenviar(j).amount
                    End If
114             Next j
            
116             cantidadTotalItem = cantidadTotalItem + ObjAEnviar.amount
            
118             If Not TieneObjetos(ObjAEnviar.ObjIndex, cantidadTotalItem, UserIndex) Then
120                 Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad disponible para agregar.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                'Si es un item recorro todo el array para ver si ese elemento ya está agregado y de paso me guardo la primer posición vacía
                Dim i As Long
122             For i = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
                    'Si encuentro el item y tiene lugar pongo Found en la posición que lo encontré
124                 If UserList(UserIndex).ComUsu.itemsAenviar(i).ObjIndex = ObjAEnviar.ObjIndex And UserList(UserIndex).ComUsu.itemsAenviar(i).amount <= 10000 Then
                        'Me fijo si le va a entrar el objeto con las cantidades en el slot que encontró
126                     If UserList(UserIndex).ComUsu.itemsAenviar(i).amount + ObjAEnviar.amount <= 10000 Then
                            'Si le entra simplemente le agrego las cantidades
128                         UserList(UserIndex).ComUsu.itemsAenviar(i).amount = UserList(UserIndex).ComUsu.itemsAenviar(i).amount + ObjAEnviar.amount
130                         nada = True
                            Exit For
                        'Si no le entra la cantidad en ese slot me guardo la posición y mas adelante me fijo si hay otra posición libre.
                        Else
132                         FoundPos = i
                        End If
                    'Si no encuentra item en la pos y todavía no guardó ninguna primera posición me la guardo.
134                 ElseIf UserList(UserIndex).ComUsu.itemsAenviar(i).ObjIndex = 0 And FirstEmptyPos = 0 Then
136                     FirstEmptyPos = i
                    End If
                
138             Next i
            
140             With UserList(UserIndex).ComUsu
                    'Si tengo una posición encontrada con un item y a su ves 1 slot vacío para agregar los restantes de ese item
142                 If FoundPos > 0 And FirstEmptyPos > 0 Then
                        Dim restante As Long
144                     restante = .itemsAenviar(FoundPos).amount + ObjAEnviar.amount - 10000
146                     If FoundPos > FirstEmptyPos Then
148                         .itemsAenviar(FoundPos).amount = restante
150                         .itemsAenviar(FirstEmptyPos).amount = 10000
                        Else
152                         .itemsAenviar(FoundPos).amount = 10000
154                         .itemsAenviar(FirstEmptyPos).amount = restante
                        End If
156                     .itemsAenviar(FirstEmptyPos).ObjIndex = ObjAEnviar.ObjIndex
158                 ElseIf FoundPos = 0 And FirstEmptyPos <> 0 Then
                        'Si entré aca es porque tengo que guardar el item en la pos vacía que encontré
160                     .itemsAenviar(FirstEmptyPos).ObjIndex = ObjAEnviar.ObjIndex
162                     .itemsAenviar(FirstEmptyPos).amount = ObjAEnviar.amount
164                 ElseIf FirstEmptyPos = 0 And nada = False Then
                        'le aviso que no le entran los items
166                     Call WriteConsoleMsg(UserIndex, "No tienes suficiente lugar para agregar esa cantidad o item", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End With
            End If
        
        
            'Le envío la data al cliente para agregar en la lista.
        
168         Call WriteChangeUserTradeSlot(AQuien, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, False)
170         Call WriteChangeUserTradeSlot(UserIndex, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, True)
        
            Exit Sub

EnviarObjetoTransaccion_Err:
172         Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.EnviarObjetoTransaccion", Erl)

        
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer, Optional ByVal Invalido As Boolean = False)
        
        On Error GoTo FinComerciarUsu_Err
        
100     If UserIndex = 0 Then Exit Sub
        

102     With UserList(UserIndex)

104         If IsValidUserRef(.ComUsu.DestUsu) And Not Invalido Then
106             Call WriteUserCommerceEnd(UserIndex)
            End If
        
108         .ComUsu.Acepto = False
110         .ComUsu.cant = 0
112         Call SetUserRef(.ComUsu.DestUsu, 0)
114         .ComUsu.Objeto = 0
116         .ComUsu.DestNick = vbNullString
118         .flags.Comerciando = False

        End With

        
        Exit Sub

FinComerciarUsu_Err:
120     Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.FinComerciarUsu", Erl)

        
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
        On Error GoTo AceptarComercioUsu_Err
        
        Dim objOfrecido As t_Obj
        Dim OtroUserIndex As Integer
        Dim TerminarAhora As Boolean

100     TerminarAhora = UserList(userIndex).ComUsu.DestUsu.ArrayIndex <= 0 Or UserList(userIndex).ComUsu.DestUsu.ArrayIndex > MaxUsers
102     OtroUserIndex = UserList(userIndex).ComUsu.DestUsu.ArrayIndex

104     If Not TerminarAhora Then
106         TerminarAhora = Not UserList(OtroUserIndex).flags.UserLogged Or Not UserList(UserIndex).flags.UserLogged
        End If

108     If Not TerminarAhora Then
110         TerminarAhora = UserList(OtroUserIndex).ComUsu.DestUsu.ArrayIndex <> userIndex
        End If

112     If TerminarAhora Then
114         Call FinComerciarUsu(UserIndex)
    
116         If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
118             Call FinComerciarUsu(OtroUserIndex)
            End If
    
            Exit Sub

        End If

120     UserList(UserIndex).ComUsu.Acepto = True

122     If UserList(OtroUserIndex).ComUsu.Acepto = False Then
124         Call WriteConsoleMsg(UserIndex, "El otro usuario aun no ha aceptado tu oferta.", e_FontTypeNames.FONTTYPE_TALK)
            Exit Sub

        End If

126     If UserList(UserIndex).ComUsu.Oro > UserList(UserIndex).Stats.GLD Then
128         Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", e_FontTypeNames.FONTTYPE_TALK)
130         TerminarAhora = True
        End If
    
132     If UserList(OtroUserIndex).ComUsu.Oro > UserList(OtroUserIndex).Stats.GLD Then
134         Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", e_FontTypeNames.FONTTYPE_TALK)
136         GoTo FinalizarComercio
        End If

        ' Verificamos que si tiene los objetos JUSTO ANTES de intercambiarlos
        Dim i As Long
138     For i = 1 To UBound(UserList(OtroUserIndex).ComUsu.itemsAenviar)
140         objOfrecido = UserList(OtroUserIndex).ComUsu.itemsAenviar(i)
142         If objOfrecido.ObjIndex > 0 And Not TieneObjetos(objOfrecido.ObjIndex, objOfrecido.amount, OtroUserIndex) Then
144             Call WriteConsoleMsg(OtroUserIndex, "El otro usuario no tiene esa cantidad disponible para ofrecer.", e_FontTypeNames.FONTTYPE_INFO)
146             GoTo FinalizarComercio
            End If
        
148         objOfrecido = UserList(UserIndex).ComUsu.itemsAenviar(i)
150         If objOfrecido.ObjIndex > 0 And Not TieneObjetos(objOfrecido.ObjIndex, objOfrecido.amount, UserIndex) Then
152             Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad disponible para ofrecer.", e_FontTypeNames.FONTTYPE_INFO)
154             GoTo FinalizarComercio
            End If
156     Next i

        'Por si las moscas...
158     If TerminarAhora Then GoTo FinalizarComercio
    
        'pone el oro directamente en la billetera
160     If UserList(OtroUserIndex).ComUsu.Oro > 0 Then
162         UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Oro
164         Call WriteUpdateUserStats(OtroUserIndex)
166         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Oro
168         Call WriteUpdateUserStats(UserIndex)
        End If
    
170     If UserList(UserIndex).ComUsu.Oro > 0 Then
172         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Oro
174         Call WriteUpdateUserStats(UserIndex)
176         UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Oro
178         Call WriteUpdateUserStats(OtroUserIndex)
        End If
        
        ' Confirmamos que SI tienen los objetos a comerciar, procedemos con el cambio.
180     For i = 1 To UBound(UserList(OtroUserIndex).ComUsu.itemsAenviar)
182         If Not MeterItemEnInventario(UserIndex, UserList(OtroUserIndex).ComUsu.itemsAenviar(i)) Then
184             Call TirarItemAlPiso(UserList(UserIndex).Pos, UserList(OtroUserIndex).ComUsu.itemsAenviar(i))
            End If
        
186         Call QuitarObjetos(UserList(OtroUserIndex).ComUsu.itemsAenviar(i).ObjIndex, UserList(OtroUserIndex).ComUsu.itemsAenviar(i).amount, OtroUserIndex)
188     Next i
    
        Dim j As Long
190     For j = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
192         If MeterItemEnInventario(OtroUserIndex, UserList(UserIndex).ComUsu.itemsAenviar(j)) = False Then
194             Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, UserList(UserIndex).ComUsu.itemsAenviar(j))
            End If
196         Call QuitarObjetos(UserList(UserIndex).ComUsu.itemsAenviar(j).ObjIndex, UserList(UserIndex).ComUsu.itemsAenviar(j).amount, UserIndex)
198     Next j


200     Call UpdateUserInv(True, UserIndex, 0)
202     Call UpdateUserInv(True, OtroUserIndex, 0)

FinalizarComercio:
204     Call FinComerciarUsu(UserIndex)
206     Call FinComerciarUsu(OtroUserIndex)
    
        Exit Sub

AceptarComercioUsu_Err:
208     Call TraceError(Err.Number, Err.Description, "mdlCOmercioConUsuario.AceptarComercioUsu", Erl)

        
End Sub
