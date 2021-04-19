Attribute VB_Name = "mdlCOmercioConUsuario"
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
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)

        On Error GoTo ErrHandler

108     If MapInfo(UserList(Origen).Pos.Map).Seguro = 0 Then
110         Call WriteConsoleMsg(Origen, "No se puede usar el comercio seguro en zona insegura.", FontTypeNames.FONTTYPE_INFO)
112         Call WriteWorkRequestTarget(Origen, 0)
            Exit Sub

        End If

        'Si ambos pusieron /comerciar entonces
114     If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
            'Actualiza el inventario del usuario
116         Call UpdateUserInv(True, Origen, 0)
            'Decirle al origen que abra la ventanita.
118         Call WriteUserCommerceInit(Origen)
120         UserList(Origen).flags.Comerciando = True

            'Actualiza el inventario del usuario
122         Call UpdateUserInv(True, Destino, 0)
            'Decirle al origen que abra la ventanita.
124         Call WriteUserCommerceInit(Destino)
126         UserList(Destino).flags.Comerciando = True
            'Limpio los arrays antes de iniciar el comercio seguro.
            Erase UserList(Origen).ComUsu.itemsAenviar
            Erase UserList(Destino).ComUsu.itemsAenviar
            UserList(Destino).ComUsu.Oro = 0
            UserList(Origen).ComUsu.Oro = 0
            
            'Call EnviarObjetoTransaccion(Origen)
        Else
            'Es el primero que comercia ?
            'Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
128         UserList(Destino).flags.TargetUser = Origen
    
130         UserList(Destino).flags.pregunta = 4
132         Call WritePreguntaBox(Destino, UserList(Origen).name & " desea comerciar contigo. ¿Aceptás?")
    
        End If

    

        Exit Sub
ErrHandler:
134     Call LogError("Error en IniciarComercioConUsuario: " & Err.Description)

End Sub
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer, ByVal UserIndex As Integer, ByRef ObjAEnviar As obj)
        
        On Error GoTo EnviarObjetoTransaccion_Err
        
        Dim FirstEmptyPos As Byte
        Dim FoundPos As Byte
        Dim nada As Boolean
        Dim cantidadTotalItem As Long
        'Me fijo si recibe oro
            If ObjAEnviar.ObjIndex = 0 Then
                'Si es oro simplemente me fijo si ya había agregado antes y se lo sumo
                If UserList(UserIndex).ComUsu.Oro + ObjAEnviar.amount <= UserList(UserIndex).Stats.GLD Then
                    UserList(UserIndex).ComUsu.Oro = UserList(UserIndex).ComUsu.Oro + ObjAEnviar.amount
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad disponible para agregar.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
            
                Dim j As Byte
                'me fijo si tiene esas cantiaddeds para que no duplique items
                For j = j To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
                    If UserList(UserIndex).ComUsu.itemsAenviar(j).ObjIndex = ObjAEnviar.ObjIndex Then
                        cantidadTotalItem = cantidadTotalItem + UserList(UserIndex).ComUsu.itemsAenviar(j).amount
                    End If
                Next j
                
                cantidadTotalItem = cantidadTotalItem + ObjAEnviar.amount
                
                If Not TieneObjetos(ObjAEnviar.ObjIndex, cantidadTotalItem, UserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad disponible para agregar.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Si es un item recorro todo el array para ver si ese elemento ya está agregado y de paso me guardo la primer posición vacía
                Dim i As Byte
                For i = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
                    'Si encuentro el item y tiene lugar pongo Found en la posición que lo encontré
                    If UserList(UserIndex).ComUsu.itemsAenviar(i).ObjIndex = ObjAEnviar.ObjIndex And UserList(UserIndex).ComUsu.itemsAenviar(i).amount < 10000 Then
                        'Me fijo si le va a entrar el objeto con las cantidades en el slot que encontró
                        If UserList(UserIndex).ComUsu.itemsAenviar(i).amount + ObjAEnviar.amount <= 10000 Then
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
                    ElseIf FirstEmptyPos = 0 And nada = False Then
                        'le aviso que no le entran los items
                        Call WriteConsoleMsg(UserIndex, "No tienes suficiente lugar para agregar esa cantidad o item", FontTypeNames.FONTTYPE_INFO)
                    End If
                End With
            End If
            
            
            'Le envío la data al cliente para agregar en la lista.
            
112         Call WriteChangeUserTradeSlot(AQuien, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, False)
            Call WriteChangeUserTradeSlot(UserIndex, UserList(UserIndex).ComUsu.itemsAenviar, UserList(UserIndex).ComUsu.Oro, True)
        Exit Sub

EnviarObjetoTransaccion_Err:
114     Call RegistrarError(Err.Number, Err.Description, "mdlCOmercioConUsuario.EnviarObjetoTransaccion", Erl)
116     Resume Next
        
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
        
        On Error GoTo FinComerciarUsu_Err
        
100     If UserIndex = 0 Then Exit Sub
        

102     With UserList(UserIndex)

104         If .ComUsu.DestUsu > 0 Then
106             Call WriteUserCommerceEnd(UserIndex)

            End If
        
108         .ComUsu.Acepto = False
110         .ComUsu.cant = 0
112         .ComUsu.DestUsu = 0
114         .ComUsu.Objeto = 0
116         .ComUsu.DestNick = vbNullString
118         .flags.Comerciando = False

        End With

        
        Exit Sub

FinComerciarUsu_Err:
120     Call RegistrarError(Err.Number, Err.Description, "mdlCOmercioConUsuario.FinComerciarUsu", Erl)
122     Resume Next
        
End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
        
        On Error GoTo AceptarComercioUsu_Err
        

        Dim Obj1          As obj, Obj2 As obj

        Dim OtroUserIndex As Integer

        Dim TerminarAhora As Boolean

100     TerminarAhora = False

102     If UserList(UserIndex).ComUsu.DestUsu <= 0 Or UserList(UserIndex).ComUsu.DestUsu > MaxUsers Then
104         TerminarAhora = True

        End If

106     OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu

108     If Not TerminarAhora Then
110         If UserList(OtroUserIndex).flags.UserLogged = False Or UserList(UserIndex).flags.UserLogged = False Then
112             TerminarAhora = True

            End If

        End If

114     If Not TerminarAhora Then
116         If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
118             TerminarAhora = True

            End If

        End If

120     If Not TerminarAhora Then
122         If UserList(OtroUserIndex).name <> UserList(UserIndex).ComUsu.DestNick Then
124             TerminarAhora = True

            End If

        End If

126     If Not TerminarAhora Then
128         If UserList(UserIndex).name <> UserList(OtroUserIndex).ComUsu.DestNick Then
130             TerminarAhora = True

            End If

        End If

132     If TerminarAhora = True Then
134         Call FinComerciarUsu(UserIndex)
    
136         If OtroUserIndex <= 0 Or OtroUserIndex > MaxUsers Then
138             Call FinComerciarUsu(OtroUserIndex)
            End If
    
            Exit Sub

        End If

140     UserList(UserIndex).ComUsu.Acepto = True
142     TerminarAhora = False

144     If UserList(OtroUserIndex).ComUsu.Acepto = False Then
146         Call WriteConsoleMsg(UserIndex, "El otro usuario aun no ha aceptado tu oferta.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub

        End If

148     If UserList(UserIndex).ComUsu.Oro > UserList(UserIndex).Stats.GLD Then
154             Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
156             TerminarAhora = True
        End If
        
        If UserList(OtroUserIndex).ComUsu.Oro > UserList(OtroUserIndex).Stats.GLD Then
                Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                TerminarAhora = True
        End If

        'Por si las moscas...
188     If TerminarAhora = True Then
190         Call FinComerciarUsu(UserIndex)
    
192         Call FinComerciarUsu(OtroUserIndex)
        
            Exit Sub

        End If

    

        '[CORREGIDO]
        'Desde acá corregí el bug que cuando se ofrecian mas de
        '10k de oro no le llegaban al destinatario.

        'pone el oro directamente en la billetera
        If UserList(OtroUserIndex).ComUsu.Oro > 0 Then
            UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.Oro
            Call WriteUpdateUserStats(OtroUserIndex)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.Oro
202         Call WriteUpdateUserStats(UserIndex)
        End If
            Dim i As Byte
            For i = 1 To UBound(UserList(OtroUserIndex).ComUsu.itemsAenviar)
                If MeterItemEnInventario(UserIndex, UserList(OtroUserIndex).ComUsu.itemsAenviar(i)) = False Then
206                 Call TirarItemAlPiso(UserList(UserIndex).Pos, UserList(OtroUserIndex).ComUsu.itemsAenviar(i))
                End If
                Call QuitarObjetos(UserList(OtroUserIndex).ComUsu.itemsAenviar(i).ObjIndex, UserList(OtroUserIndex).ComUsu.itemsAenviar(i).amount, OtroUserIndex)
            Next i
        
        If UserList(UserIndex).ComUsu.Oro > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.Oro
            Call WriteUpdateUserStats(OtroUserIndex)
            UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.Oro
            Call WriteUpdateUserStats(OtroUserIndex)
        End If
        
            Dim j As Byte
            For j = 1 To UBound(UserList(UserIndex).ComUsu.itemsAenviar)
                If MeterItemEnInventario(OtroUserIndex, UserList(UserIndex).ComUsu.itemsAenviar(j)) = False Then
                    Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, UserList(UserIndex).ComUsu.itemsAenviar(j))
                End If
                Call QuitarObjetos(UserList(UserIndex).ComUsu.itemsAenviar(j).ObjIndex, UserList(UserIndex).ComUsu.itemsAenviar(j).amount, UserIndex)
            Next j


        '[/CORREGIDO] :p

226     Call UpdateUserInv(True, UserIndex, 0)
228     Call UpdateUserInv(True, OtroUserIndex, 0)

230     Call FinComerciarUsu(UserIndex)
232     Call FinComerciarUsu(OtroUserIndex)
 
        
        Exit Sub

AceptarComercioUsu_Err:
234     Call RegistrarError(Err.Number, Err.Description, "mdlCOmercioConUsuario.AceptarComercioUsu", Erl)
236     Resume Next
        
End Sub

'[/Alejo]
