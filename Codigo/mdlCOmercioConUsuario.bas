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

Public Type tCOmercioUsuario

    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto As Integer 'Indice del inventario a comerciar, que objeto desea dar
    
    'El tipo de datos de Cant ahora es Long (antes Integer)
    'asi se puede comerciar con oro > 32k
    '[CORREGIDO]
    cant As Long 'Cuantos comerciar, cuantos objetos desea dar
    '[/CORREGIDO]
    Acepto As Boolean

End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)

    On Error GoTo Errhandler

    If UserList(Origen).flags.BattleModo = 1 Then
        Call WriteConsoleMsg(Origen, "No podes usar el sistema de comercio en el battle.", FontTypeNames.FONTTYPE_EXP)
        Exit Sub

    End If

    If UserList(Destino).flags.BattleModo = 1 Then
        Call WriteConsoleMsg(Destino, "No podes usar el sistema de comercio en el battle.", FontTypeNames.FONTTYPE_EXP)
        Exit Sub

    End If

    If MapInfo(UserList(Origen).Pos.Map).Seguro = 0 Then
        Call WriteConsoleMsg(Origen, "No se puede usar el comercio seguro en zona insegura.", FontTypeNames.FONTTYPE_INFO)
        Call WriteWorkRequestTarget(Origen, 0)
        Exit Sub

    End If

    'Si ambos pusieron /comerciar entonces
    If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
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

        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        'Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, Escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.TargetUser = Origen
    
        UserList(Destino).flags.pregunta = 4
        Call WritePreguntaBox(Destino, UserList(Origen).name & " desea comerciar contigo. �Acept�s?")
    
    End If

    

    Exit Sub
Errhandler:
    Call LogError("Error en IniciarComercioConUsuario: " & Err.description)

End Sub

'envia a AQuien el objeto del otro
Public Sub EnviarObjetoTransaccion(ByVal AQuien As Integer)
        
        On Error GoTo EnviarObjetoTransaccion_Err
        

        Dim ObjInd  As Integer

        Dim ObjCant As Long

        '[Alejo]: En esta funcion se centralizaba el problema
        '         de no poder comerciar con mas de 32k de oro.
        '         Ahora si funciona!!!

100     ObjCant = UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.cant

102     If UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto = FLAGORO Then
104         ObjInd = iORO
        Else
106         ObjInd = UserList(UserList(AQuien).ComUsu.DestUsu).Invent.Object(UserList(UserList(AQuien).ComUsu.DestUsu).ComUsu.Objeto).ObjIndex

        End If

108     If ObjCant <= 0 Or ObjInd <= 0 Then Exit Sub

110     If ObjInd > 0 And ObjCant > 0 Then
112         Call WriteChangeUserTradeSlot(AQuien, ObjInd, ObjCant)
        

        End If

        
        Exit Sub

EnviarObjetoTransaccion_Err:
        Call RegistrarError(Err.Number, Err.description, "mdlCOmercioConUsuario.EnviarObjetoTransaccion", Erl)
        Resume Next
        
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
        
        On Error GoTo FinComerciarUsu_Err
        

100     With UserList(UserIndex)

102         If .ComUsu.DestUsu > 0 Then
104             Call WriteUserCommerceEnd(UserIndex)

            End If
        
106         .ComUsu.Acepto = False
108         .ComUsu.cant = 0
110         .ComUsu.DestUsu = 0
112         .ComUsu.Objeto = 0
114         .ComUsu.DestNick = vbNullString
116         .flags.Comerciando = False

        End With

        
        Exit Sub

FinComerciarUsu_Err:
        Call RegistrarError(Err.Number, Err.description, "mdlCOmercioConUsuario.FinComerciarUsu", Erl)
        Resume Next
        
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

148     If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
150         Obj1.ObjIndex = iORO

152         If UserList(UserIndex).ComUsu.cant > UserList(UserIndex).Stats.GLD Then
154             Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
156             TerminarAhora = True

            End If

        Else
158         Obj1.Amount = UserList(UserIndex).ComUsu.cant
160         Obj1.ObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).ObjIndex

162         If Obj1.Amount > UserList(UserIndex).Invent.Object(UserList(UserIndex).ComUsu.Objeto).Amount Then
164             Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
166             TerminarAhora = True

            End If

        End If

168     If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
170         Obj2.ObjIndex = iORO

172         If UserList(OtroUserIndex).ComUsu.cant > UserList(OtroUserIndex).Stats.GLD Then
174             Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
176             TerminarAhora = True

            End If

        Else
178         Obj2.Amount = UserList(OtroUserIndex).ComUsu.cant
180         Obj2.ObjIndex = UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).ObjIndex

182         If Obj2.Amount > UserList(OtroUserIndex).Invent.Object(UserList(OtroUserIndex).ComUsu.Objeto).Amount Then
184             Call WriteConsoleMsg(OtroUserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
186             TerminarAhora = True

            End If

        End If

        'Por si las moscas...
188     If TerminarAhora = True Then
190         Call FinComerciarUsu(UserIndex)
    
192         Call FinComerciarUsu(OtroUserIndex)
        
            Exit Sub

        End If

    

        '[CORREGIDO]
        'Desde ac� correg� el bug que cuando se ofrecian mas de
        '10k de oro no le llegaban al destinatario.

        'pone el oro directamente en la billetera
194     If UserList(OtroUserIndex).ComUsu.Objeto = FLAGORO Then
            'quito la cantidad de oro ofrecida
196         UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD - UserList(OtroUserIndex).ComUsu.cant
            ' If UserList(OtroUserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " solto oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.cant)
198         Call WriteUpdateUserStats(OtroUserIndex)
            'y se la doy al otro
200         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(OtroUserIndex).ComUsu.cant
            'If UserList(OtroUserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(UserIndex).name & " recibio oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(OtroUserIndex).ComUsu.cant)
            'Esta linea del log es al pedo.
202         Call WriteUpdateUserStats(UserIndex)
        Else

            'Quita el objeto y se lo da al otro
204         If MeterItemEnInventario(UserIndex, Obj2) = False Then
206             Call TirarItemAlPiso(UserList(UserIndex).Pos, Obj2)

            End If

208         Call QuitarObjetos(Obj2.ObjIndex, Obj2.Amount, OtroUserIndex)
    
            'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
            'If ObjData(Obj2.ObjIndex).Log = 1 Then
            '  '   Call LogDesarrollo(UserList(OtroUserIndex).name & " le pas� en comercio seguro a " & UserList(UserIndex).name & " " & Obj2.Amount & " " & ObjData(Obj2.ObjIndex).name)
            ' End If
            'Es mucha cantidad?
            ' If Obj2.Amount > MAX_OBJ_LOGUEABLE Then
            'Si no es de los prohibidos de loguear, lo logueamos.
            ' If ObjData(Obj2.ObjIndex).NoLog <> 1 Then
            '    Call LogDesarrollo(UserList(OtroUserIndex).name & " le pas� en comercio seguro a " & UserList(UserIndex).name & " " & Obj2.Amount & " " & ObjData(Obj2.ObjIndex).name)
            ' End If
            ' End If
        End If

        'pone el oro directamente en la billetera
210     If UserList(UserIndex).ComUsu.Objeto = FLAGORO Then
            'quito la cantidad de oro ofrecida
212         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - UserList(UserIndex).ComUsu.cant
            ' If UserList(UserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(UserIndex).name & " solt� oro en comercio seguro con " & UserList(OtroUserIndex).name & ". Cantidad: " & UserList(UserIndex).ComUsu.cant)
214         Call WriteUpdateUserStats(UserIndex)
            'y se la doy al otro
216         UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + UserList(UserIndex).ComUsu.cant
            'If UserList(UserIndex).ComUsu.cant > MAX_ORO_LOGUEABLE Then Call LogDesarrollo(UserList(OtroUserIndex).name & " recibio oro en comercio seguro con " & UserList(UserIndex).name & ". Cantidad: " & UserList(UserIndex).ComUsu.cant)
            'Esta linea del log es al pedo.
218         Call WriteUpdateUserStats(OtroUserIndex)
        Else

            'Quita el objeto y se lo da al otro
220         If MeterItemEnInventario(OtroUserIndex, Obj1) = False Then
222             Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, Obj1)

            End If

224         Call QuitarObjetos(Obj1.ObjIndex, Obj1.Amount, UserIndex)
    
            'Es un Objeto que tenemos que loguear? Pablo (ToxicWaste) 07/09/07
            ' If ObjData(Obj1.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " le pas� en comercio seguro a " & UserList(OtroUserIndex).name & " " & Obj1.Amount & " " & ObjData(Obj1.ObjIndex).name)
            ' End If
            'Es mucha cantidad?
            ' If Obj1.Amount > MAX_OBJ_LOGUEABLE Then
            'Si no es de los prohibidos de loguear, lo logueamos.
            '  If ObjData(Obj1.ObjIndex).NoLog <> 1 Then
            ''     Call LogDesarrollo(UserList(OtroUserIndex).name & " le pas� en comercio seguro a " & UserList(UserIndex).name & " " & Obj1.Amount & " " & ObjData(Obj1.ObjIndex).name)
            '  End If
            ' End If
    
        End If

        '[/CORREGIDO] :p

226     Call UpdateUserInv(True, UserIndex, 0)
228     Call UpdateUserInv(True, OtroUserIndex, 0)

230     Call FinComerciarUsu(UserIndex)
232     Call FinComerciarUsu(OtroUserIndex)
 
        
        Exit Sub

AceptarComercioUsu_Err:
        Call RegistrarError(Err.Number, Err.description, "mdlCOmercioConUsuario.AceptarComercioUsu", Erl)
        Resume Next
        
End Sub

'[/Alejo]
