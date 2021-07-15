Attribute VB_Name = "InvUsuario"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
        On Error GoTo TieneObjetosRobables_Err
      

        Dim i        As Integer
        Dim ObjIndex As Integer

100     For i = 1 To UserList(UserIndex).CurrentInventorySlots
102         ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

104         If ObjIndex > 0 Then
106             If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And ObjData(ObjIndex).OBJType <> eOBJType.OtDonador And ObjData(ObjIndex).OBJType <> eOBJType.otRunas) Then
108                 TieneObjetosRobables = True
                    Exit Function

                End If
    
            End If

110     Next i

        
        Exit Function

TieneObjetosRobables_Err:
112     Call TraceError(Err.Number, Err.Description, "InvUsuario.TieneObjetosRobables", Erl)

        
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean

        On Error GoTo manejador

        Dim flag As Boolean

100     If Slot <> 0 Then
102         If UserList(UserIndex).Invent.Object(Slot).Equipped Then
104             ClasePuedeUsarItem = True
                Exit Function

            End If

        End If

106     If EsGM(UserIndex) Then
108         ClasePuedeUsarItem = True
            Exit Function

        End If

        Dim i As Integer

110     For i = 1 To NUMCLASES

112         If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
114             ClasePuedeUsarItem = False
                Exit Function

            End If

116     Next i

118     ClasePuedeUsarItem = True

        Exit Function

manejador:
120     LogError ("Error en ClasePuedeUsarItem")

End Function

Function RazaPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean
        On Error GoTo RazaPuedeUsarItem_Err

        Dim Objeto As ObjData, i As Long
        
100     Objeto = ObjData(ObjIndex)
        
102     If EsGM(UserIndex) Then
104         RazaPuedeUsarItem = True
            Exit Function
        End If

106     For i = 1 To NUMRAZAS
108         If Objeto.RazaProhibida(i) = UserList(UserIndex).raza Then
110             RazaPuedeUsarItem = False
                Exit Function
            End If

112     Next i
        
        ' Si el objeto no define una raza en particular
114     If Objeto.RazaDrow + Objeto.RazaElfa + Objeto.RazaEnana + Objeto.RazaGnoma + Objeto.RazaHumana + Objeto.RazaOrca = 0 Then
116         RazaPuedeUsarItem = True
        
        Else ' El objeto esta definido para alguna raza en especial
118         Select Case UserList(UserIndex).raza
                Case eRaza.Humano
120                 RazaPuedeUsarItem = Objeto.RazaHumana > 0

122             Case eRaza.Elfo
124                 RazaPuedeUsarItem = Objeto.RazaElfa > 0
                
126             Case eRaza.Drow
128                 RazaPuedeUsarItem = Objeto.RazaDrow > 0
    
130             Case eRaza.Orco
132                 RazaPuedeUsarItem = Objeto.RazaOrca > 0
                    
134             Case eRaza.Gnomo
136                 RazaPuedeUsarItem = Objeto.RazaGnoma > 0
                    
138             Case eRaza.Enano
140                 RazaPuedeUsarItem = Objeto.RazaEnana > 0

            End Select
        End If
        
        Exit Function

RazaPuedeUsarItem_Err:
142     LogError ("Error en RazaPuedeUsarItem")

End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
        
        On Error GoTo QuitarNewbieObj_Err
        

        Dim j As Integer

100     For j = 1 To UserList(UserIndex).CurrentInventorySlots

102         If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
104             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then
106                 Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
108                 Call UpdateUserInv(False, UserIndex, j)

                End If
        
            End If

110     Next j
    
        'Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
        'es transportado a su hogar de origen ;)
112     If MapInfo(UserList(UserIndex).Pos.Map).Newbie Then
        
            Dim DeDonde As WorldPos
        
114         Select Case UserList(UserIndex).Hogar

                Case eCiudad.cUllathorpe
116                 DeDonde = Ullathorpe

118             Case eCiudad.cNix
120                 DeDonde = Nix
    
122             Case eCiudad.cBanderbill
124                 DeDonde = Banderbill
            
126             Case eCiudad.cLindos
128                 DeDonde = Lindos
                
130             Case eCiudad.cArghal
132                 DeDonde = Arghal
                
134             Case eCiudad.cArkhein
136                 DeDonde = Arkhein
                
138             Case Else
140                 DeDonde = Ullathorpe

            End Select
        
142         Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
        End If

        
        Exit Sub

QuitarNewbieObj_Err:
144     Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarNewbieObj", Erl)

        
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
        
        On Error GoTo LimpiarInventario_Err
        

        Dim j As Integer

100     For j = 1 To UserList(UserIndex).CurrentInventorySlots
102         UserList(UserIndex).Invent.Object(j).ObjIndex = 0
104         UserList(UserIndex).Invent.Object(j).amount = 0
106         UserList(UserIndex).Invent.Object(j).Equipped = 0
        
        Next

108     UserList(UserIndex).Invent.NroItems = 0

110     UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
112     UserList(UserIndex).Invent.ArmourEqpSlot = 0

114     UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
116     UserList(UserIndex).Invent.WeaponEqpSlot = 0

118     UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
120     UserList(UserIndex).Invent.HerramientaEqpSlot = 0

122     UserList(UserIndex).Invent.CascoEqpObjIndex = 0
124     UserList(UserIndex).Invent.CascoEqpSlot = 0

126     UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
128     UserList(UserIndex).Invent.EscudoEqpSlot = 0

130     UserList(UserIndex).Invent.DañoMagicoEqpObjIndex = 0
132     UserList(UserIndex).Invent.DañoMagicoEqpSlot = 0

134     UserList(UserIndex).Invent.ResistenciaEqpObjIndex = 0
136     UserList(UserIndex).Invent.ResistenciaEqpSlot = 0

138     UserList(UserIndex).Invent.NudilloObjIndex = 0
140     UserList(UserIndex).Invent.NudilloSlot = 0

142     UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
144     UserList(UserIndex).Invent.MunicionEqpSlot = 0

146     UserList(UserIndex).Invent.BarcoObjIndex = 0
148     UserList(UserIndex).Invent.BarcoSlot = 0

150     UserList(UserIndex).Invent.MonturaObjIndex = 0
152     UserList(UserIndex).Invent.MonturaSlot = 0

154     UserList(UserIndex).Invent.MagicoObjIndex = 0
156     UserList(UserIndex).Invent.MagicoSlot = 0

        
        Exit Sub

LimpiarInventario_Err:
158     Call TraceError(Err.Number, Err.Description, "InvUsuario.LimpiarInventario", Erl)

        
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
        '***************************************************
        On Error GoTo ErrHandler
        
100     With UserList(UserIndex)
        
            ' GM's (excepto Dioses y Admins) no pueden tirar oro
102         If (.flags.Privilegios And (PlayerType.user Or PlayerType.Admin Or PlayerType.Dios)) = 0 Then
104             Call LogGM(.Name, " trató de tirar " & PonerPuntos(Cantidad) & " de oro en " & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
                Exit Sub

            End If
         
            ' Si el usuario tiene ORO, entonces lo tiramos
106         If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then

                Dim i     As Byte
                Dim MiObj As obj

                'info debug
                Dim loops As Integer
                Dim Extra    As Long
                Dim TeniaOro As Long

108             TeniaOro = .Stats.GLD

110             If Cantidad > 100000 Then 'Para evitar explotar demasiado
112                 Extra = Cantidad - 100000
114                 Cantidad = 100000

                End If
        
116             Do While (Cantidad > 0)
            
118                 If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
120                     MiObj.amount = MAX_INVENTORY_OBJS
122                     Cantidad = Cantidad - MiObj.amount
                    Else
124                     MiObj.amount = Cantidad
126                     Cantidad = Cantidad - MiObj.amount

                    End If

128                 MiObj.ObjIndex = iORO

                    Dim AuxPos As WorldPos

130                 If .clase = eClass.Pirat Then
132                     AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    Else
134                     AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    End If
            
136                 If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
138                     .Stats.GLD = .Stats.GLD - MiObj.amount

                    End If
            
                    'info debug
140                 loops = loops + 1

142                 If loops > 100 Then
144                     Call LogError("Se ha superado el limite de iteraciones(100) permitido en el Sub TirarOro()")
                        Exit Sub

                    End If
            
                Loop
                
                ' Si es GM, registramos lo q hizo incluso si es Horacio
146             If EsGM(UserIndex) Then

148                 If MiObj.ObjIndex = iORO Then
150                     Call LogGM(.Name, "Tiro: " & PonerPuntos(Cantidad) & " monedas de oro.")

                    Else
152                     Call LogGM(.Name, "Tiro cantidad:" & PonerPuntos(Cantidad) & " Objeto:" & ObjData(MiObj.ObjIndex).Name)

                    End If

                End If
        
154             If TeniaOro = .Stats.GLD Then Extra = 0

156             If Extra > 0 Then
158                 .Stats.GLD = .Stats.GLD - Extra
                End If
    
160             Call WriteUpdateGold(UserIndex)
            End If
        
        End With

        Exit Sub

ErrHandler:
162 Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarOro", Erl())
    
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarUserInvItem_Err
        

100     If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
102     With UserList(UserIndex).Invent.Object(Slot)

104         If .amount <= Cantidad And .Equipped = 1 Then
106             Call Desequipar(UserIndex, Slot)

            End If
        
            'Quita un objeto
108         .amount = .amount - Cantidad

            '¿Quedan mas?
110         If .amount <= 0 Then
112             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
114             .ObjIndex = 0
116             .amount = 0

            End If

        End With

        
        Exit Sub

QuitarUserInvItem_Err:
118     Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarUserInvItem", Erl)

        
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
        
        On Error GoTo UpdateUserInv_Err
        

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
104             Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
            Else
106             Call ChangeUserInv(UserIndex, Slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots

                'Actualiza el inventario
110             If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
112                 Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
                Else
114                 Call ChangeUserInv(UserIndex, LoopC, NullObj)

                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateUserInv_Err:
118     Call TraceError(Err.Number, Err.Description, "InvUsuario.UpdateUserInv", Erl)

        
End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal Slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        
        On Error GoTo DropObj_Err

        Dim obj As obj

100     If num > 0 Then
            
102         With UserList(UserIndex)

104             If num > .Invent.Object(Slot).amount Then
106                 num = .Invent.Object(Slot).amount
                End If
    
108             obj.ObjIndex = .Invent.Object(Slot).ObjIndex
110             obj.amount = num
    
112             If ObjData(obj.ObjIndex).Destruye = 0 Then
    
                    'Check objeto en el suelo
114                 If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Or num + MapData(.Pos.Map, X, Y).ObjInfo.amount <= MAX_INVENTORY_OBJS Then
                      
116                     If num + MapData(.Pos.Map, X, Y).ObjInfo.amount > MAX_INVENTORY_OBJS Then
118                         num = MAX_INVENTORY_OBJS - MapData(.Pos.Map, X, Y).ObjInfo.amount
                        End If
                        
                        ' Si sos Admin, Dios o Usuario, crea el objeto en el piso.
120                     If (.flags.Privilegios And (PlayerType.user Or PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

                            ' Tiramos el item al piso
122                         Call MakeObj(obj, Map, X, Y)

                        End If
                        
124                     Call QuitarUserInvItem(UserIndex, Slot, num)
126                     Call UpdateUserInv(False, UserIndex, Slot)
                        
128                     If Not .flags.Privilegios And PlayerType.user Then
130                         Call LogGM(.Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name)
                        End If
    
                    Else
                    
                        'Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
132                     Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                Else
134                 Call QuitarUserInvItem(UserIndex, Slot, num)
136                 Call UpdateUserInv(False, UserIndex, Slot)
    
                End If
            
            End With

        End If
        
        Exit Sub

DropObj_Err:
138     Call TraceError(Err.Number, Err.Description, "InvUsuario.DropObj", Erl)


        
End Sub

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo EraseObj_Err
        

        Dim Rango As Byte

100     MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - num

102     If MapData(Map, X, Y).ObjInfo.amount <= 0 Then

104         If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType <> otTeleport Then
106             Call QuitarItemLimpieza(Map, X, Y)
            End If
            
108         MapData(Map, X, Y).ObjInfo.ObjIndex = 0
110         MapData(Map, X, Y).ObjInfo.amount = 0
    
    
112         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

        
        Exit Sub

EraseObj_Err:
114     Call TraceError(Err.Number, Err.Description, "InvUsuario.EraseObj", Erl)

        
End Sub

Sub MakeObj(ByRef obj As obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Limpiar As Boolean = True)
        
        On Error GoTo MakeObj_Err

        Dim Color As Long

        Dim Rango As Byte

100     If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
102         If MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex Then
104             MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount + obj.amount
            Else
                ' Lo agrego a la limpieza del mundo o reseteo el timer si el objeto ya existía
106             If ObjData(obj.ObjIndex).OBJType <> otTeleport Then
108                 Call AgregarItemLimpieza(Map, X, Y, MapData(Map, X, Y).ObjInfo.ObjIndex <> 0)
                End If
            
110             MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex

112             If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
114                 MapData(Map, X, Y).ObjInfo.amount = ObjData(obj.ObjIndex).VidaUtil
                Else
116                 MapData(Map, X, Y).ObjInfo.amount = obj.amount

                End If
                
            End If
            
118         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(obj.ObjIndex, MapData(Map, X, Y).ObjInfo.amount, X, Y))
    
        End If
        
        Exit Sub

MakeObj_Err:
120     Call TraceError(Err.Number, Err.Description, "InvUsuario.MakeObj", Erl)


        
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean

        On Error GoTo ErrHandler

        'Call LogTarea("MeterItemEnInventario")
 
        Dim X    As Integer

        Dim Y    As Integer

        Dim Slot As Byte

        '¿el user ya tiene un objeto del mismo tipo? ?????
100     Slot = 1

102     Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS
104         Slot = Slot + 1

106         If Slot > UserList(UserIndex).CurrentInventorySlots Then
                Exit Do

            End If

        Loop
        
        'Sino busca un slot vacio
108     If Slot > UserList(UserIndex).CurrentInventorySlots Then
110         Slot = 1

112         Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
114             Slot = Slot + 1

116             If Slot > UserList(UserIndex).CurrentInventorySlots Then
                    'Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
118                 Call WriteLocaleMsg(UserIndex, "328", FontTypeNames.FONTTYPE_FIGHT)
120                 MeterItemEnInventario = False
                    Exit Function

                End If

            Loop
122         UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

        End If
        
        'Mete el objeto
124     If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
126         UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
128         UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
        Else
130         UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS

        End If
        
132     Call UpdateUserInv(False, UserIndex, Slot)

134     MeterItemEnInventario = True

        Exit Function
ErrHandler:

End Function

Function HayLugarEnInventario(ByVal UserIndex As Integer) As Boolean

        On Error GoTo ErrHandler
 
        Dim X    As Integer

        Dim Y    As Integer

        Dim Slot As Byte

100     Slot = 1

102     Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
104         Slot = Slot + 1
106         If Slot > UserList(UserIndex).CurrentInventorySlots Then
108             HayLugarEnInventario = False
                Exit Function
            End If
        Loop
        
110     HayLugarEnInventario = True

        Exit Function
ErrHandler:

End Function

Sub GetObj(ByVal UserIndex As Integer)
        
        On Error GoTo GetObj_Err
        
        Dim X    As Integer
        Dim Y    As Integer
        Dim Slot As Byte
        Dim obj   As ObjData
        Dim MiObj As obj

        '¿Hay algun obj?
100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
102         If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

104             If UserList(UserIndex).flags.Montado = 1 Then
106                 Call WriteConsoleMsg(UserIndex, "Debes descender de tu montura para agarrar objetos del suelo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
108             X = UserList(UserIndex).Pos.X
110             Y = UserList(UserIndex).Pos.Y
112             obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
114             MiObj.amount = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.amount
116             MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex
        
118             If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Else
            
                    'Quitamos el objeto
120                 Call EraseObj(MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

122                 If Not UserList(UserIndex).flags.Privilegios And PlayerType.user Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
    
124                 If BusquedaTesoroActiva Then
126                     If UserList(UserIndex).Pos.Map = TesoroNumMapa And UserList(UserIndex).Pos.X = TesoroX And UserList(UserIndex).Pos.Y = TesoroY Then
    
128                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).Name & " encontro el tesoro ¡Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
130                         BusquedaTesoroActiva = False

                        End If

                    End If
                
132                 If BusquedaRegaloActiva Then
134                     If UserList(UserIndex).Pos.Map = RegaloNumMapa And UserList(UserIndex).Pos.X = RegaloX And UserList(UserIndex).Pos.Y = RegaloY Then
136                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).Name & " fue el valiente que encontro el gran item magico ¡Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
138                         BusquedaRegaloActiva = False

                        End If

                    End If
                
                    'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                    'Es un Objeto que tenemos que loguear?
140                 If ObjData(MiObj.ObjIndex).Log = 1 Then
142                     Call LogDesarrollo(UserList(UserIndex).Name & " juntó del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name)

                        ' ElseIf MiObj.Amount = 1000 Then 'Es mucha cantidad?
                        '  'Si no es de los prohibidos de loguear, lo logueamos.
                        '   'If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                        ' Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
                        ' End If
                    End If
                
                End If

            End If

        Else

144         If Not UserList(UserIndex).flags.UltimoMensaje = 261 Then
146             Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)
148             UserList(UserIndex).flags.UltimoMensaje = 261

            End If
    
            'Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
        End If

        
        Exit Sub

GetObj_Err:
150     Call TraceError(Err.Number, Err.Description, "InvUsuario.GetObj", Erl)

        
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
        
        On Error GoTo Desequipar_Err

        'Desequipa el item slot del inventario
        Dim obj As ObjData

100     If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
            Exit Sub
102     ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
            Exit Sub

        End If

104     obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

106     Select Case obj.OBJType

            Case eOBJType.otWeapon
108             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
110             UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
112             UserList(UserIndex).Invent.WeaponEqpSlot = 0
114             UserList(UserIndex).Char.Arma_Aura = ""
116             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 1))
        
118             UserList(UserIndex).Char.WeaponAnim = NingunArma
            
120             If UserList(UserIndex).flags.Montado = 0 Then
122                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                
124             If obj.MagicDamageBonus > 0 Then
126                 Call WriteUpdateDM(UserIndex)
                End If
    
128         Case eOBJType.otFlechas
130             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
132             UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
134             UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
                ' Case eOBJType.otAnillos
                '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
                '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
                ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
            
136         Case eOBJType.otHerramientas
138             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
140             UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
142             UserList(UserIndex).Invent.HerramientaEqpSlot = 0

144             If UserList(UserIndex).flags.UsandoMacro = True Then
146                 Call WriteMacroTrabajoToggle(UserIndex, False)
                End If
        
148             UserList(UserIndex).Char.WeaponAnim = NingunArma
            
150             If UserList(UserIndex).flags.Montado = 0 Then
152                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
       
154         Case eOBJType.otMagicos
    
156             Select Case obj.EfectoMagico

                    Case 1 'Regenera Energia
158                     UserList(UserIndex).flags.RegeneracionSta = 0

160                 Case 2 'Modifica los Atributos
                        If obj.QueAtributo <> 0 Then
162                         UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
164                         UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                            ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                            
166                         Call WriteFYA(UserIndex)
                        End If

168                 Case 3 'Modifica los skills
                        If obj.QueSkill <> 0 Then
170                         UserList(UserIndex).Stats.UserSkills(obj.QueSkill) = UserList(UserIndex).Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento
                        End If
                        
172                 Case 4 ' Regeneracion Vida
174                     UserList(UserIndex).flags.RegeneracionHP = 0

176                 Case 5 ' Regeneracion Mana
178                     UserList(UserIndex).flags.RegeneracionMana = 0

180                 Case 6 'Aumento Golpe
182                     UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit - obj.CuantoAumento
184                     UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - obj.CuantoAumento

186                 Case 7 '
                
188                 Case 9 ' Orbe Ignea
190                     UserList(UserIndex).flags.NoMagiaEfecto = 0

192                 Case 10
194                     UserList(UserIndex).flags.incinera = 0

196                 Case 11
198                     UserList(UserIndex).flags.Paraliza = 0

200                 Case 12

202                     If UserList(UserIndex).flags.Muerto = 0 Then UserList(UserIndex).flags.CarroMineria = 0
                
204                 Case 14
                        'UserList(UserIndex).flags.DañoMagico = 0
                
206                 Case 15 'Pendiete del Sacrificio
208                     UserList(UserIndex).flags.PendienteDelSacrificio = 0
                 
210                 Case 16
212                     UserList(UserIndex).flags.NoPalabrasMagicas = 0

214                 Case 17 'Sortija de la verdad
216                     UserList(UserIndex).flags.NoDetectable = 0

218                 Case 18 ' Pendiente del Experto
220                     UserList(UserIndex).flags.PendienteDelExperto = 0

222                 Case 19 ' Envenenamiento
224                     UserList(UserIndex).flags.Envenena = 0

226                 Case 20 ' anillo de las sombras
228                     UserList(UserIndex).flags.AnilloOcultismo = 0
                
                End Select
        
230             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 5))
232             UserList(UserIndex).Char.Otra_Aura = 0
234             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
236             UserList(UserIndex).Invent.MagicoObjIndex = 0
238             UserList(UserIndex).Invent.MagicoSlot = 0
        
240         Case eOBJType.otNudillos
    
                'falta mandar animacion
            
242             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
244             UserList(UserIndex).Invent.NudilloObjIndex = 0
246             UserList(UserIndex).Invent.NudilloSlot = 0
        
248             UserList(UserIndex).Char.Arma_Aura = ""
250             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 1))
        
252             UserList(UserIndex).Char.WeaponAnim = NingunArma
254             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        
256         Case eOBJType.otArmadura
258             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
260             UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
262             UserList(UserIndex).Invent.ArmourEqpSlot = 0
        
264             If UserList(UserIndex).flags.Navegando = 0 Then
266                 If UserList(UserIndex).flags.Montado = 0 Then
268                     Call DarCuerpoDesnudo(UserIndex)
270                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                End If
        
272             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 2))
        
274             UserList(UserIndex).Char.Body_Aura = 0

276             If obj.ResistenciaMagica > 0 Then
278                 Call WriteUpdateRM(UserIndex)
                End If
    
280         Case eOBJType.otCasco
282             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
284             UserList(UserIndex).Invent.CascoEqpObjIndex = 0
286             UserList(UserIndex).Invent.CascoEqpSlot = 0
288             UserList(UserIndex).Char.Head_Aura = 0
290             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 4))

292             UserList(UserIndex).Char.CascoAnim = NingunCasco
294             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    
296             If obj.ResistenciaMagica > 0 Then
298                 Call WriteUpdateRM(UserIndex)
                End If
    
300         Case eOBJType.otEscudo
302             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
304             UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
306             UserList(UserIndex).Invent.EscudoEqpSlot = 0
308             UserList(UserIndex).Char.Escudo_Aura = 0
310             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 3))
        
312             UserList(UserIndex).Char.ShieldAnim = NingunEscudo

314             If UserList(UserIndex).flags.Montado = 0 Then
316                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                End If
                
318             If obj.ResistenciaMagica > 0 Then
320                 Call WriteUpdateRM(UserIndex)
                End If
                
322         Case eOBJType.otDañoMagico
324             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
326             UserList(UserIndex).Invent.DañoMagicoEqpObjIndex = 0
328             UserList(UserIndex).Invent.DañoMagicoEqpSlot = 0
330             UserList(UserIndex).Char.DM_Aura = 0
332             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 6))
334             Call WriteUpdateDM(UserIndex)
                
336         Case eOBJType.otResistencia
338             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
340             UserList(UserIndex).Invent.ResistenciaEqpObjIndex = 0
342             UserList(UserIndex).Invent.ResistenciaEqpSlot = 0
344             UserList(UserIndex).Char.RM_Aura = 0
346             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 7))
348             Call WriteUpdateRM(UserIndex)
        
        End Select

350     Call UpdateUserInv(False, UserIndex, Slot)

        
        Exit Sub

Desequipar_Err:
352     Call TraceError(Err.Number, Err.Description, "InvUsuario.Desequipar", Erl)

        
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then
102         SexoPuedeUsarItem = True
            Exit Function

        End If

104     If ObjData(ObjIndex).Mujer = 1 Then
106         SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
108     ElseIf ObjData(ObjIndex).Hombre = 1 Then
110         SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
        Else
112         SexoPuedeUsarItem = True

        End If

        Exit Function
ErrHandler:
114     Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
        
        On Error GoTo FaccionPuedeUsarItem_Err
        
100     If EsGM(UserIndex) Then
102         FaccionPuedeUsarItem = True
            Exit Function
        End If
        
104     If ObjIndex < 1 Then Exit Function

106     If ObjData(ObjIndex).Real = 1 Then
108         If Status(UserIndex) = 3 Then
110             FaccionPuedeUsarItem = esArmada(UserIndex)
            Else
112             FaccionPuedeUsarItem = False

            End If

114     ElseIf ObjData(ObjIndex).Caos = 1 Then

116         If Status(UserIndex) = 2 Then
118             FaccionPuedeUsarItem = esCaos(UserIndex)
            Else
120             FaccionPuedeUsarItem = False

            End If

        Else
122         FaccionPuedeUsarItem = True

        End If

        
        Exit Function

FaccionPuedeUsarItem_Err:
124     Call TraceError(Err.Number, Err.Description, "InvUsuario.FaccionPuedeUsarItem", Erl)

        
End Function

'Equipa barco y hace el cambio de ropaje correspondiente
Sub EquiparBarco(ByVal UserIndex As Integer)
      On Error GoTo EquiparBarco_Err

      Dim Barco As ObjData

100   With UserList(UserIndex)
102     Barco = ObjData(.Invent.BarcoObjIndex)

104     If .flags.Muerto = 1 Then
106       If Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw Then
              ' No tenemos la cabeza copada que va con iRopaBuceoMuerto,
              ' asique asignamos el casper directamente caminando sobre el agua.
108           .Char.Body = iCuerpoMuerto 'iRopaBuceoMuerto
110           .Char.Head = iCabezaMuerto
          ElseIf Barco.Ropaje = iTrajeAltoNw Then
          
          ElseIf Barco.Ropaje = iTrajeBajoNw Then
          
          Else
112           .Char.Body = iFragataFantasmal
114           .Char.Head = 0
          End If
      
        Else ' Esta vivo
116       If Barco.Ropaje = iTraje Then
118         .Char.Body = iTraje
120         .Char.Head = .OrigChar.Head
122         If .Invent.CascoEqpObjIndex > 0 Then
124           .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            End If
          ElseIf Barco.Ropaje = iTrajeAltoNw Then
            .Char.Body = iTrajeAltoNw
            .Char.Head = .OrigChar.Head
            If .Invent.CascoEqpObjIndex > 0 Then
                .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            End If
          ElseIf Barco.Ropaje = iTrajeBajoNw Then
            .Char.Body = iTrajeBajoNw
            .Char.Head = .OrigChar.Head
            If .Invent.CascoEqpObjIndex > 0 Then
                .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
            End If
          Else
126         .Char.Head = 0
128         .Char.CascoAnim = NingunCasco
          End If

130       If .Faccion.ArmadaReal = 1 Then
132         If Barco.Ropaje = iBarca Then .Char.Body = iBarcaArmada
134         If Barco.Ropaje = iGalera Then .Char.Body = iGaleraArmada
136         If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonArmada

138       ElseIf .Faccion.FuerzasCaos = 1 Then
140         If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCaos
142         If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCaos
144         If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCaos
          
          Else
146         If Barco.Ropaje = iBarca Then .Char.Body = IIf(.Faccion.Status = 0, iBarcaCrimi, iBarcaCiuda)
148         If Barco.Ropaje = iGalera Then .Char.Body = IIf(.Faccion.Status = 0, iGaleraCrimi, iGaleraCiuda)
150         If Barco.Ropaje = iGaleon Then .Char.Body = IIf(.Faccion.Status = 0, iGaleonCrimi, iGaleonCiuda)
          End If
        End If

152     .Char.ShieldAnim = NingunEscudo
154     .Char.WeaponAnim = NingunArma
    
156     Call WriteNadarToggle(UserIndex, (Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw), (Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw))
      End With
  
      Exit Sub

EquiparBarco_Err:
158   Call TraceError(Err.Number, Err.Description, "InvUsuario.EquiparBarco", Erl)


End Sub

'Equipa un item del inventario
Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
        On Error GoTo ErrHandler

        Dim obj       As ObjData
        Dim ObjIndex  As Integer
        Dim errordesc As String

100     ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
102     obj = ObjData(ObjIndex)
        
104     If PuedeUsarObjeto(UserIndex, ObjIndex, True) > 0 Then
            Exit Sub
        End If

106     With UserList(UserIndex)
108          If .flags.Muerto = 1 Then
                 'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
110              Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                 Exit Sub

             End If

112         Select Case obj.OBJType
                Case eOBJType.otWeapon
114                 errordesc = "Arma"

                    'Si esta equipado lo quita
116                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
118                     Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
120                     .Char.WeaponAnim = NingunArma

122                     If .flags.Montado = 0 Then
124                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If

                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
126                 If .Invent.WeaponEqpObjIndex > 0 Then
128                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
            
130                 If .Invent.HerramientaEqpObjIndex > 0 Then
132                     Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
                    End If
            
134                 If .Invent.NudilloObjIndex > 0 Then
136                     Call Desequipar(UserIndex, .Invent.NudilloSlot)
                    End If

138                 .Invent.Object(Slot).Equipped = 1
140                 .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
142                 .Invent.WeaponEqpSlot = Slot
            
144                 If obj.Proyectil = 1 Then 'Si es un arco, desequipa el escudo.

146                     If .Invent.EscudoEqpObjIndex = 1700 Or _
                           .Invent.EscudoEqpObjIndex = 1730 Or _
                           .Invent.EscudoEqpObjIndex = 1724 Or _
                           .Invent.EscudoEqpObjIndex = 1717 Or _
                           .Invent.EscudoEqpObjIndex = 1699 Then
                           ' Estos escudos SI pueden ser usados con arco.
                        Else

148                         If .Invent.EscudoEqpObjIndex > 0 Then
150                             Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
152                             Call WriteConsoleMsg(UserIndex, "No podes tirar flechas si tenés un escudo equipado. Tu escudo fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        End If

                    End If

154                 If obj.DosManos = 1 Then
156                     If .Invent.EscudoEqpObjIndex > 0 Then
158                         Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
160                         Call WriteConsoleMsg(UserIndex, "No puedes usar armas dos manos si tienes un escudo equipado. Tu escudo fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    End If

                    'Sonido
162                 If obj.SndAura = 0 Then
164                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
166                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If

168                 If Len(obj.CreaGRH) <> 0 Then
170                     .Char.Arma_Aura = obj.CreaGRH
172                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If

174                 If obj.MagicDamageBonus > 0 Then
176                     Call WriteUpdateDM(UserIndex)
                    End If
                
178                 If .flags.Montado = 0 Then
                
180                     If .flags.Navegando = 0 Then
182                         .Char.WeaponAnim = obj.WeaponAnim
184                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
      
186             Case eOBJType.otHerramientas

                    'Si esta equipado lo quita
188                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
190                     Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
192                 If .Invent.HerramientaEqpObjIndex > 0 Then
194                     Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
                    End If
             
196                 If .Invent.WeaponEqpObjIndex > 0 Then
198                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
             
200                 .Invent.Object(Slot).Equipped = 1
202                 .Invent.HerramientaEqpObjIndex = ObjIndex
204                 .Invent.HerramientaEqpSlot = Slot
             
206                 If .flags.Montado = 0 Then
                
208                     If .flags.Navegando = 0 Then
210                         .Char.WeaponAnim = obj.WeaponAnim
212                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
       
214             Case eOBJType.otMagicos
216                 errordesc = "Magico"

                    'Si esta equipado lo quita
218                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
220                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If

                    'Quitamos el elemento anterior
222                 If .Invent.MagicoObjIndex > 0 Then
224                     Call Desequipar(UserIndex, .Invent.MagicoSlot)
                    End If
        
226                 .Invent.Object(Slot).Equipped = 1
228                 .Invent.MagicoObjIndex = .Invent.Object(Slot).ObjIndex
230                 .Invent.MagicoSlot = Slot
                
                    ' Debug.Print "magico" & obj.EfectoMagico
232                 Select Case obj.EfectoMagico
                        Case 1 ' Regenera Stamina
234                         .flags.RegeneracionSta = 1

236                     Case 2 'Modif la fuerza, agilidad, carisma, etc
                            ' .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo)
238                         .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
240                         .Stats.UserAtributos(obj.QueAtributo) = MinimoInt(.Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento, .Stats.UserAtributosBackUP(obj.QueAtributo) * 2)
                
242                         Call WriteFYA(UserIndex)

244                     Case 3 'Modifica los skills
            
246                         .Stats.UserSkills(obj.QueSkill) = .Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento

248                     Case 4
250                         .flags.RegeneracionHP = 1

252                     Case 5
254                         .flags.RegeneracionMana = 1

256                     Case 6
258                         .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
260                         .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento

262                     Case 9
264                         .flags.NoMagiaEfecto = 1

266                     Case 10
268                         .flags.incinera = 1

270                     Case 11
272                         .flags.Paraliza = 1

274                     Case 12
276                         .flags.CarroMineria = 1
                
278                     Case 14
                            '.flags.DañoMagico = obj.CuantoAumento
                
280                     Case 15 'Pendiete del Sacrificio
282                         .flags.PendienteDelSacrificio = 1

284                     Case 16
286                         .flags.NoPalabrasMagicas = 1

288                     Case 17
290                         .flags.NoDetectable = 1
                   
292                     Case 18 ' Pendiente del Experto
294                         .flags.PendienteDelExperto = 1

296                     Case 19
298                         .flags.Envenena = 1

300                     Case 20 'Anillo ocultismo
302                         .flags.AnilloOcultismo = 1
    
                    End Select
            
                    'Sonido
304                 If obj.SndAura <> 0 Then
306                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
            
308                 If Len(obj.CreaGRH) <> 0 Then
310                     .Char.Otra_Aura = obj.CreaGRH
312                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
                    End If
        
                    'Call WriteUpdateExp(UserIndex)
                    'Call CheckUserLevel(UserIndex)
            
314             Case eOBJType.otNudillos
316                 If .Invent.WeaponEqpObjIndex > 0 Then
318                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                    End If

320                 If .Invent.Object(Slot).Equipped Then
322                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If

                    'Quitamos el elemento anterior
324                 If .Invent.NudilloObjIndex > 0 Then
326                     Call Desequipar(UserIndex, .Invent.NudilloSlot)

                    End If

328                 .Invent.Object(Slot).Equipped = 1
330                 .Invent.NudilloObjIndex = .Invent.Object(Slot).ObjIndex
332                 .Invent.NudilloSlot = Slot

                    'Falta enviar anim
334                 If .flags.Montado = 0 Then
                
336                     If .flags.Navegando = 0 Then
338                         .Char.WeaponAnim = obj.WeaponAnim
340                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If

342                 If obj.SndAura = 0 Then
344                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
346                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
                 
348                 If Len(obj.CreaGRH) <> 0 Then
350                     .Char.Arma_Aura = obj.CreaGRH
352                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If
    
354             Case eOBJType.otFlechas
                    'Si esta equipado lo quita
356                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
358                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
360                 If .Invent.MunicionEqpObjIndex > 0 Then
362                     Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                    End If
        
364                 .Invent.Object(Slot).Equipped = 1
366                 .Invent.MunicionEqpObjIndex = .Invent.Object(Slot).ObjIndex
368                 .Invent.MunicionEqpSlot = Slot

370             Case eOBJType.otArmadura
372                 If obj.Ropaje = 0 Then
374                     Call WriteConsoleMsg(UserIndex, "Hay un error con este objeto. Infórmale a un administrador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Si esta equipado lo quita
376                 If .Invent.Object(Slot).Equipped Then
378                     Call Desequipar(UserIndex, Slot)

380                     If .flags.Navegando = 0 And .flags.Montado = 0 Then
382                         Call DarCuerpoDesnudo(UserIndex)
384                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Else
386                         .flags.Desnudo = 1
                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
388                 If .Invent.ArmourEqpObjIndex > 0 Then
390                     errordesc = "Armadura 2"
392                     Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
394                     errordesc = "Armadura 3"

                    End If
  
                    'Lo equipa
396                 If Len(obj.CreaGRH) <> 0 Then
398                     .Char.Body_Aura = obj.CreaGRH
400                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))

                    End If
            
402                 .Invent.Object(Slot).Equipped = 1
404                 .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
406                 .Invent.ArmourEqpSlot = Slot

408                 If .flags.Montado = 0 And .flags.Navegando = 0 Then
410                     .Char.Body = obj.Ropaje

412                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    
414                 .flags.Desnudo = 0

416                 If obj.ResistenciaMagica > 0 Then
418                     Call WriteUpdateRM(UserIndex)
                    End If
    
420             Case eOBJType.otCasco
                    'Si esta equipado lo quita
422                 If .Invent.Object(Slot).Equipped Then
424                     Call Desequipar(UserIndex, Slot)
                
426                     .Char.CascoAnim = NingunCasco
428                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Exit Sub

                    End If
    
                    'Quita el anterior
430                 If .Invent.CascoEqpObjIndex > 0 Then
432                     Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If

434                 errordesc = "Casco"

                    'Lo equipa
436                 If Len(obj.CreaGRH) <> 0 Then
438                     .Char.Head_Aura = obj.CreaGRH
440                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
                    End If
            
442                 .Invent.Object(Slot).Equipped = 1
444                 .Invent.CascoEqpObjIndex = .Invent.Object(Slot).ObjIndex
446                 .Invent.CascoEqpSlot = Slot
            
448                 If .flags.Navegando = 0 Then
450                     .Char.CascoAnim = obj.CascoAnim
452                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                
454                 If obj.ResistenciaMagica > 0 Then
456                     Call WriteUpdateRM(UserIndex)
                    End If

458             Case eOBJType.otEscudo
                    'Si esta equipado lo quita
460                 If .Invent.Object(Slot).Equipped Then
462                     Call Desequipar(UserIndex, Slot)
                 
464                     .Char.ShieldAnim = NingunEscudo

466                     If .flags.Montado = 0 And .flags.Navegando = 0 Then
468                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
     
                    'Quita el anterior
470                 If .Invent.EscudoEqpObjIndex > 0 Then
472                     Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                    End If
     
                    'Lo equipa
474                 If .Invent.Object(Slot).ObjIndex = 1700 Or _
                       .Invent.Object(Slot).ObjIndex = 1730 Or _
                       .Invent.Object(Slot).ObjIndex = 1724 Or _
                       .Invent.Object(Slot).ObjIndex = 1717 Or _
                       .Invent.Object(Slot).ObjIndex = 1699 Then
             
                    Else

476                     If .Invent.WeaponEqpObjIndex > 0 Then
478                         If ObjData(.Invent.WeaponEqpObjIndex).Proyectil = 1 Then
480                             Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
482                             Call WriteConsoleMsg(UserIndex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                            End If
                        End If

                    End If

484                 If .Invent.WeaponEqpObjIndex > 0 Then
486                     If ObjData(.Invent.WeaponEqpObjIndex).DosManos = 1 Then
488                         Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
490                         Call WriteConsoleMsg(UserIndex, "No puedes equipar un escudo si tienes un arma dos manos equipada. Tu arma fue desequipada.", FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    End If

492                 errordesc = "Escudo"

494                 If Len(obj.CreaGRH) <> 0 Then
496                     .Char.Escudo_Aura = obj.CreaGRH
498                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
                    End If

500                 .Invent.Object(Slot).Equipped = 1
502                 .Invent.EscudoEqpObjIndex = .Invent.Object(Slot).ObjIndex
504                 .Invent.EscudoEqpSlot = Slot

506                 If .flags.Navegando = 0 And .flags.Montado = 0 Then
508                     .Char.ShieldAnim = obj.ShieldAnim
510                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If

512                 If obj.ResistenciaMagica > 0 Then
514                     Call WriteUpdateRM(UserIndex)
                    End If

516             Case eOBJType.otDañoMagico
                    'Si esta equipado lo quita
518                 If .Invent.Object(Slot).Equipped Then
520                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
522                 If .Invent.DañoMagicoEqpSlot > 0 Then
524                     Call Desequipar(UserIndex, .Invent.DañoMagicoEqpSlot)
                    End If

526                 .Invent.Object(Slot).Equipped = 1
528                 .Invent.DañoMagicoEqpObjIndex = .Invent.Object(Slot).ObjIndex
530                 .Invent.DañoMagicoEqpSlot = Slot

532                 If Len(obj.CreaGRH) <> 0 Then
534                     .Char.DM_Aura = obj.CreaGRH
536                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.DM_Aura, False, 6))
                    End If

538                 Call WriteUpdateDM(UserIndex)

540             Case eOBJType.otResistencia
                    'Si esta equipado lo quita
542                 If .Invent.Object(Slot).Equipped Then
544                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
546                 If .Invent.ResistenciaEqpSlot > 0 Then
548                     Call Desequipar(UserIndex, .Invent.ResistenciaEqpSlot)
                    End If
                
550                 .Invent.Object(Slot).Equipped = 1
552                 .Invent.ResistenciaEqpObjIndex = .Invent.Object(Slot).ObjIndex
554                 .Invent.ResistenciaEqpSlot = Slot
                
556                 If Len(obj.CreaGRH) <> 0 Then
558                     .Char.RM_Aura = obj.CreaGRH
560                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.RM_Aura, False, 7))
                    End If

562                 Call WriteUpdateRM(UserIndex)

            End Select

        End With

        'Actualiza
564     Call UpdateUserInv(False, UserIndex, Slot)

        Exit Sub
    
ErrHandler:
566     Debug.Print errordesc
568     Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description & "- " & errordesc)

End Sub

Public Function CheckClaseTipo(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then

102         CheckClaseTipo = True
            Exit Function

        End If

104     Select Case ObjData(ItemIndex).ClaseTipo

            Case 0
106             CheckClaseTipo = True
                Exit Function

108         Case 2

110             If UserList(UserIndex).clase = eClass.Mage Then CheckClaseTipo = True
112             If UserList(UserIndex).clase = eClass.Druid Then CheckClaseTipo = True
                Exit Function

114         Case 1

116             If UserList(UserIndex).clase = eClass.Warrior Then CheckClaseTipo = True
118             If UserList(UserIndex).clase = eClass.Assasin Then CheckClaseTipo = True
120             If UserList(UserIndex).clase = eClass.Bard Then CheckClaseTipo = True
122             If UserList(UserIndex).clase = eClass.Cleric Then CheckClaseTipo = True
124             If UserList(UserIndex).clase = eClass.Paladin Then CheckClaseTipo = True
126             If UserList(UserIndex).clase = eClass.Trabajador Then CheckClaseTipo = True
128             If UserList(UserIndex).clase = eClass.Hunter Then CheckClaseTipo = True
                Exit Function

        End Select

        Exit Function
ErrHandler:
130     Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo hErr

        '*************************************************
        'Author: Unknown
        'Last modified: 24/01/2007
        'Handels the usage of items from inventory box.
        '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
        '24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
        '*************************************************

        Dim obj      As ObjData

        Dim ObjIndex As Integer

        Dim TargObj  As ObjData

        Dim MiObj    As obj
        
100     With UserList(UserIndex)

102         If .Invent.Object(Slot).amount = 0 Then Exit Sub
    
104         obj = ObjData(.Invent.Object(Slot).ObjIndex)
    
106         If obj.OBJType = eOBJType.otWeapon Then
108             If obj.Proyectil = 1 Then
    
                    'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
110                 If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
                Else
    
                    'dagas
112                 If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    
                End If
    
            Else
    
114             If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
116             If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then Exit Sub
    
            End If
    
118         If .flags.Meditando Then
120             .flags.Meditando = False
122             .Char.FX = 0
124             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
            End If
    
126         If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
128             Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
    
130         If .Stats.ELV < obj.MinELV Then
132             Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
    
134         ObjIndex = .Invent.Object(Slot).ObjIndex
136         .flags.TargetObjInvIndex = ObjIndex
138         .flags.TargetObjInvSlot = Slot
    
140         Select Case obj.OBJType
    
                Case eOBJType.otUseOnce
    
142                 If .flags.Muerto = 1 Then
144                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
                    'Usa el item
146                 .Stats.MinHam = .Stats.MinHam + obj.MinHam
    
148                 If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
150                 .flags.Hambre = 0
152                 Call WriteUpdateHungerAndThirst(UserIndex)
                    'Sonido

154                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, .Pos.X, .Pos.Y))

                    'Quitamos del inv el item
156                 Call QuitarUserInvItem(UserIndex, Slot, 1)
            
158                 Call UpdateUserInv(False, UserIndex, Slot)
    
160             Case eOBJType.otGuita
    
162                 If .flags.Muerto = 1 Then
164                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
166                 .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).amount
168                 .Invent.Object(Slot).amount = 0
170                 .Invent.Object(Slot).ObjIndex = 0
172                 .Invent.NroItems = .Invent.NroItems - 1
            
174                 Call UpdateUserInv(False, UserIndex, Slot)
176                 Call WriteUpdateGold(UserIndex)
            
178             Case eOBJType.otWeapon
    
180                 If .flags.Muerto = 1 Then
182                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
184                 If Not .Stats.MinSta > 0 Then
186                     Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
188                 If ObjData(ObjIndex).Proyectil = 1 Then
                        'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
190                     Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                    Else
192                     If .flags.TargetObj = Leña Then
194                         If .Invent.Object(Slot).ObjIndex = DAGA Then
196                             Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                            End If
                        End If
    
                    End If
            
                    'REVISAR LADDER
                    'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
198                 If .Invent.Object(Slot).Equipped = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
200             Case eOBJType.otHerramientas
    
202                 If .flags.Muerto = 1 Then
204                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
206                 If Not .Stats.MinSta > 0 Then
208                     Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
                    'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
210                 If .Invent.Object(Slot).Equipped = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
212                     Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
214                 Select Case obj.Subtipo
                    
                        Case 1, 2  ' Herramientas del Pescador - Caña y Red
216                         Call WriteWorkRequestTarget(UserIndex, eSkill.Pescar)
                    
218                     Case 3     ' Herramientas de Alquimia - Tijeras
220                         Call WriteWorkRequestTarget(UserIndex, eSkill.Alquimia)
                    
222                     Case 4     ' Herramientas de Alquimia - Olla
224                         Call EnivarObjConstruiblesAlquimia(UserIndex)
226                         Call WriteShowAlquimiaForm(UserIndex)
                    
228                     Case 5     ' Herramientas de Carpinteria - Serrucho
230                         Call EnivarObjConstruibles(UserIndex)
232                         Call WriteShowCarpenterForm(UserIndex)
                    
234                     Case 6     ' Herramientas de Tala - Hacha
236                         Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)
    
238                     Case 7     ' Herramientas de Herrero - Martillo
240                         Call WriteConsoleMsg(UserIndex, "Debes hacer click derecho sobre el yunque.", FontTypeNames.FONTTYPE_INFOIAO)
    
242                     Case 8     ' Herramientas de Mineria - Piquete
244                         Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
                    
246                     Case 9     ' Herramientas de Sastreria - Costurero
248                         Call EnivarObjConstruiblesSastre(UserIndex)
250                         Call WriteShowSastreForm(UserIndex)
    
                    End Select
        
252             Case eOBJType.otPociones
    
254                 If .flags.Muerto = 1 Then
256                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
258                 .flags.TomoPocion = True
260                 .flags.TipoPocion = obj.TipoPocion
                    
                    Dim CabezaFinal  As Integer
    
                    Dim CabezaActual As Integer
    
262                 Select Case .flags.TipoPocion
            
                        Case 1 'Modif la agilidad
264                         .flags.DuracionEfecto = obj.DuracionEfecto
            
                            'Usa el item
266                         .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(.Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)
                    
268                         Call WriteFYA(UserIndex)
                    
                            'Quitamos del inv el item
270                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
272                         If obj.Snd1 <> 0 Then
274                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
276                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
            
278                     Case 2 'Modif la fuerza
280                         .flags.DuracionEfecto = obj.DuracionEfecto
            
                            'Usa el item
282                         .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(.Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)
                    
                            'Quitamos del inv el item
284                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
286                         If obj.Snd1 <> 0 Then
288                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
290                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
292                         Call WriteFYA(UserIndex)
    
294                     Case 3 'Pocion roja, restaura HP
                    
                            'Usa el item
296                         .Stats.MinHp = .Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)
    
298                         If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                    
                            'Quitamos del inv el item
300                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
302                         If obj.Snd1 <> 0 Then
304                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                            Else
306                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
                
308                     Case 4 'Pocion azul, restaura MANA
                
                            Dim porcentajeRec As Byte
310                         porcentajeRec = obj.Porcentaje
                    
                            'Usa el item
312                         .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, porcentajeRec)
    
314                         If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                    
                            'Quitamos del inv el item
316                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
318                         If obj.Snd1 <> 0 Then
320                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                            Else
322                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
                    
324                     Case 5 ' Pocion violeta
    
326                         If .flags.Envenenado > 0 Then
328                             .flags.Envenenado = 0
330                             Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                                'Quitamos del inv el item
332                             Call QuitarUserInvItem(UserIndex, Slot, 1)
    
334                             If obj.Snd1 <> 0 Then
336                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                                Else
338                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                                End If
    
                            Else
340                             Call WriteConsoleMsg(UserIndex, "¡No te encuentras envenenado!", FontTypeNames.FONTTYPE_INFO)
    
                            End If
                    
342                     Case 6  ' Remueve Parálisis
    
344                         If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
346                             If .flags.Paralizado = 1 Then
348                                 .flags.Paralizado = 0
350                                 Call WriteParalizeOK(UserIndex)
    
                                End If
                            
352                             If .flags.Inmovilizado = 1 Then
354                                 .Counters.Inmovilizado = 0
356                                 .flags.Inmovilizado = 0
358                                 Call WriteInmovilizaOK(UserIndex)
    
                                End If
                            
                            
                            
360                             Call QuitarUserInvItem(UserIndex, Slot, 1)
    
362                             If obj.Snd1 <> 0 Then
364                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                                Else
366                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(255, .Pos.X, .Pos.Y))
    
                                End If
    
368                             Call WriteConsoleMsg(UserIndex, "Te has removido la paralizis.", FontTypeNames.FONTTYPE_INFOIAO)
                            Else
370                             Call WriteConsoleMsg(UserIndex, "No estas paralizado.", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
                    
372                     Case 7  ' Pocion Naranja
374                         .Stats.MinSta = .Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)
    
376                         If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                        
                            'Quitamos del inv el item
378                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
380                         If obj.Snd1 <> 0 Then
382                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                            Else
384                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
386                     Case 8  ' Pocion cambio cara
    
388                         Select Case .genero
    
                                Case eGenero.Hombre
    
390                                 Select Case .raza
    
                                        Case eRaza.Humano
392                                         CabezaFinal = RandomNumber(1, 40)
    
394                                     Case eRaza.Elfo
396                                         CabezaFinal = RandomNumber(101, 132)
    
398                                     Case eRaza.Drow
400                                         CabezaFinal = RandomNumber(201, 229)
    
402                                     Case eRaza.Enano
404                                         CabezaFinal = RandomNumber(301, 329)
    
406                                     Case eRaza.Gnomo
408                                         CabezaFinal = RandomNumber(401, 429)
    
410                                     Case eRaza.Orco
412                                         CabezaFinal = RandomNumber(501, 529)
    
                                    End Select
    
414                             Case eGenero.Mujer
    
416                                 Select Case .raza
    
                                        Case eRaza.Humano
418                                         CabezaFinal = RandomNumber(50, 80)
    
420                                     Case eRaza.Elfo
422                                         CabezaFinal = RandomNumber(150, 179)
    
424                                     Case eRaza.Drow
426                                         CabezaFinal = RandomNumber(250, 279)
    
428                                     Case eRaza.Gnomo
430                                         CabezaFinal = RandomNumber(350, 379)
    
432                                     Case eRaza.Enano
434                                         CabezaFinal = RandomNumber(450, 479)
    
436                                     Case eRaza.Orco
438                                         CabezaFinal = RandomNumber(550, 579)
    
                                    End Select
    
                            End Select
                
440                         .Char.Head = CabezaFinal
442                         .OrigChar.Head = CabezaFinal
444                         Call ChangeUserChar(UserIndex, .Char.Body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                            'Quitamos del inv el item
446                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 102, 0))
    
448                         If CabezaActual <> CabezaFinal Then
450                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
452                             Call WriteConsoleMsg(UserIndex, "¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
454                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
456                     Case 9  ' Pocion sexo
        
458                         Select Case .genero
    
                                Case eGenero.Hombre
460                                 .genero = eGenero.Mujer
                        
462                             Case eGenero.Mujer
464                                 .genero = eGenero.Hombre
                        
                            End Select
                
466                         Select Case .genero
    
                                Case eGenero.Hombre
    
468                                 Select Case .raza
    
                                        Case eRaza.Humano
470                                         CabezaFinal = RandomNumber(1, 40)
    
472                                     Case eRaza.Elfo
474                                         CabezaFinal = RandomNumber(101, 132)
    
476                                     Case eRaza.Drow
478                                         CabezaFinal = RandomNumber(201, 229)
    
480                                     Case eRaza.Enano
482                                         CabezaFinal = RandomNumber(301, 329)
    
484                                     Case eRaza.Gnomo
486                                         CabezaFinal = RandomNumber(401, 429)
    
488                                     Case eRaza.Orco
490                                         CabezaFinal = RandomNumber(501, 529)
    
                                    End Select
    
492                             Case eGenero.Mujer
    
494                                 Select Case .raza
    
                                        Case eRaza.Humano
496                                         CabezaFinal = RandomNumber(50, 80)
    
498                                     Case eRaza.Elfo
500                                         CabezaFinal = RandomNumber(150, 179)
    
502                                     Case eRaza.Drow
504                                         CabezaFinal = RandomNumber(250, 279)
    
506                                     Case eRaza.Gnomo
508                                         CabezaFinal = RandomNumber(350, 379)
    
510                                     Case eRaza.Enano
512                                         CabezaFinal = RandomNumber(450, 479)
    
514                                     Case eRaza.Orco
516                                         CabezaFinal = RandomNumber(550, 579)
    
                                    End Select
    
                            End Select
                
518                         .Char.Head = CabezaFinal
520                         .OrigChar.Head = CabezaFinal
522                         Call ChangeUserChar(UserIndex, .Char.Body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                            'Quitamos del inv el item
524                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 102, 0))
526                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
528                         If obj.Snd1 <> 0 Then
530                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
532                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
                    
534                     Case 10  ' Invisibilidad
                
536                         If .flags.invisible = 0 Then
538                             .flags.invisible = 1
540                             .Counters.Invisibilidad = obj.DuracionEfecto
542                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
544                             Call WriteContadores(UserIndex)
546                             Call QuitarUserInvItem(UserIndex, Slot, 1)
    
548                             If obj.Snd1 <> 0 Then
550                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                                Else
552                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("123", .Pos.X, .Pos.Y))
    
                                End If
    
554                             Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras...", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            
                            Else
556                             Call WriteConsoleMsg(UserIndex, "Ya estas invisible.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                Exit Sub
    
                            End If
                        
558                     Case 11  ' Experiencia
    
                            Dim HR   As Integer
    
                            Dim MS   As Integer
    
                            Dim SS   As Integer
    
                            Dim secs As Integer
    
560                         If .flags.ScrollExp = 1 Then
562                             .flags.ScrollExp = obj.CuantoAumento
564                             .Counters.ScrollExperiencia = obj.DuracionEfecto
566                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            
568                             secs = obj.DuracionEfecto
570                             HR = secs \ 3600
572                             MS = (secs Mod 3600) \ 60
574                             SS = (secs Mod 3600) Mod 60
    
576                             If SS > 9 Then
578                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                                Else
580                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
    
                                End If
    
                            Else
582                             Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                                Exit Sub
    
                            End If
    
584                         Call WriteContadores(UserIndex)
    
586                         If obj.Snd1 <> 0 Then
588                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
590                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
592                     Case 12  ' Oro
                
594                         If .flags.ScrollOro = 1 Then
596                             .flags.ScrollOro = obj.CuantoAumento
598                             .Counters.ScrollOro = obj.DuracionEfecto
600                             Call QuitarUserInvItem(UserIndex, Slot, 1)
602                             secs = obj.DuracionEfecto
604                             HR = secs \ 3600
606                             MS = (secs Mod 3600) \ 60
608                             SS = (secs Mod 3600) Mod 60
    
610                             If SS > 9 Then
612                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                                Else
614                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
    
                                End If
                            
                            Else
616                             Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                                Exit Sub
    
                            End If
    
618                         Call WriteContadores(UserIndex)
    
620                         If obj.Snd1 <> 0 Then
622                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
624                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
                        ' Poción que limpia todo
626                     Case 13
                    
628                         Call QuitarUserInvItem(UserIndex, Slot, 1)
630                         .flags.Envenenado = 0
632                         .flags.Incinerado = 0
                        
634                         If .flags.Inmovilizado = 1 Then
636                             .Counters.Inmovilizado = 0
638                             .flags.Inmovilizado = 0
640                             Call WriteInmovilizaOK(UserIndex)
                            
    
                            End If
                        
642                         If .flags.Paralizado = 1 Then
644                             .flags.Paralizado = 0
646                             Call WriteParalizeOK(UserIndex)
                            
    
                            End If
                        
648                         If .flags.Ceguera = 1 Then
650                             .flags.Ceguera = 0
652                             Call WriteBlindNoMore(UserIndex)
                            
    
                            End If
                        
654                         If .flags.Maldicion = 1 Then
656                             .flags.Maldicion = 0
658                             .Counters.Maldicion = 0
    
                            End If
                        
660                         .Stats.MinSta = .Stats.MaxSta
662                         .Stats.MinAGU = .Stats.MaxAGU
664                         .Stats.MinMAN = .Stats.MaxMAN
666                         .Stats.MinHp = .Stats.MaxHp
668                         .Stats.MinHam = .Stats.MaxHam
                        
670                         .flags.Hambre = 0
672                         .flags.Sed = 0
                        
674                         Call WriteUpdateHungerAndThirst(UserIndex)
676                         Call WriteConsoleMsg(UserIndex, "Donador> Te sentis sano y lleno.", FontTypeNames.FONTTYPE_WARNING)
    
678                         If obj.Snd1 <> 0 Then
680                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
682                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
                        ' Poción runa
684                     Case 14
                                       
686                         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
688                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                        
                            Dim Map     As Integer
    
                            Dim X       As Byte
    
                            Dim Y       As Byte
    
                            Dim DeDonde As WorldPos
    
690                         Call QuitarUserInvItem(UserIndex, Slot, 1)
                
692                         Select Case .Hogar
    
                                Case eCiudad.cUllathorpe
694                                 DeDonde = Ullathorpe
                                
696                             Case eCiudad.cNix
698                                 DeDonde = Nix
                    
700                             Case eCiudad.cBanderbill
702                                 DeDonde = Banderbill
                            
704                             Case eCiudad.cLindos
706                                 DeDonde = Lindos
                                
708                             Case eCiudad.cArghal
710                                 DeDonde = Arghal
                                
712                             Case eCiudad.cArkhein
714                                 DeDonde = Arkhein
                                
716                             Case Else
718                                 DeDonde = Ullathorpe
    
                            End Select
                        
720                         Map = DeDonde.Map
722                         X = DeDonde.X
724                         Y = DeDonde.Y
                        
726                         Call FindLegalPos(UserIndex, Map, X, Y)
728                         Call WarpUserChar(UserIndex, Map, X, Y, True)
730                         Call WriteConsoleMsg(UserIndex, "Ya estas a salvo...", FontTypeNames.FONTTYPE_WARNING)
    
732                         If obj.Snd1 <> 0 Then
734                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
736                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
738                     Case 15  ' Aliento de sirena
                            
740                         If .Counters.Oxigeno >= 3540 Then
                            
742                             Call WriteConsoleMsg(UserIndex, "No podes acumular más de 59 minutos de oxigeno.", FontTypeNames.FONTTYPE_INFOIAO)
744                             secs = .Counters.Oxigeno
746                             HR = secs \ 3600
748                             MS = (secs Mod 3600) \ 60
750                             SS = (secs Mod 3600) Mod 60
    
752                             If SS > 9 Then
754                                 Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                                Else
756                                 Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)
    
                                End If
    
                            Else
                                
758                             .Counters.Oxigeno = .Counters.Oxigeno + obj.DuracionEfecto
760                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                                
                                'secs = .Counters.Oxigeno
                                ' HR = secs \ 3600
                                ' MS = (secs Mod 3600) \ 60
                                ' SS = (secs Mod 3600) Mod 60
                                ' If SS > 9 Then
                                ' Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                                'Call WriteConsoleMsg(UserIndex, "Has agregado " & MS & ":" & SS & " segundos de oxigeno.", FontTypeNames.FONTTYPE_New_DONADOR)
                                ' Else
                                ' Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)
                                ' End If
                                
762                             .flags.Ahogandose = 0
764                             Call WriteOxigeno(UserIndex)
                                
766                             Call WriteContadores(UserIndex)
    
768                             If obj.Snd1 <> 0 Then
770                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                                Else
772                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                                End If
    
                            End If
    
774                     Case 16 ' Divorcio
    
776                         If .flags.Casado = 1 Then
    
                                Dim tUser As Integer
    
                                '.flags.Pareja
778                             tUser = NameIndex(.flags.Pareja)
780                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            
782                             If tUser <= 0 Then
    
                                    Dim FileUser As String
    
784                                 FileUser = CharPath & UCase$(.flags.Pareja) & ".chr"
                                    'Call WriteVar(FileUser, "FLAGS", "CASADO", 0)
                                    'Call WriteVar(FileUser, "FLAGS", "PAREJA", "")
786                                 .flags.Casado = 0
788                                 .flags.Pareja = ""
790                                 Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
792                                 .MENSAJEINFORMACION = .Name & " se ha divorciado de ti."
    
                                Else
794                                 UserList(tUser).flags.Casado = 0
796                                 UserList(tUser).flags.Pareja = ""
798                                 .flags.Casado = 0
800                                 .flags.Pareja = ""
802                                 Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
804                                 Call WriteConsoleMsg(tUser, .Name & " se ha divorciado de ti.", FontTypeNames.FONTTYPE_INFOIAO)
                                
                                End If
    
806                             If obj.Snd1 <> 0 Then
808                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                                Else
810                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                                End If
                        
                            Else
812                             Call WriteConsoleMsg(UserIndex, "No estas casado.", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
814                     Case 17 'Cara legendaria
    
816                         Select Case .genero
    
                                Case eGenero.Hombre
    
818                                 Select Case .raza
    
                                        Case eRaza.Humano
820                                         CabezaFinal = RandomNumber(684, 686)
    
822                                     Case eRaza.Elfo
824                                         CabezaFinal = RandomNumber(690, 692)
    
826                                     Case eRaza.Drow
828                                         CabezaFinal = RandomNumber(696, 698)
    
830                                     Case eRaza.Enano
832                                         CabezaFinal = RandomNumber(702, 704)
    
834                                     Case eRaza.Gnomo
836                                         CabezaFinal = RandomNumber(708, 710)
    
838                                     Case eRaza.Orco
840                                         CabezaFinal = RandomNumber(714, 716)
    
                                    End Select
    
842                             Case eGenero.Mujer
    
844                                 Select Case .raza
    
                                        Case eRaza.Humano
846                                         CabezaFinal = RandomNumber(687, 689)
    
848                                     Case eRaza.Elfo
850                                         CabezaFinal = RandomNumber(693, 695)
    
852                                     Case eRaza.Drow
854                                         CabezaFinal = RandomNumber(699, 701)
    
856                                     Case eRaza.Gnomo
858                                         CabezaFinal = RandomNumber(705, 707)
    
860                                     Case eRaza.Enano
862                                         CabezaFinal = RandomNumber(711, 713)
    
864                                     Case eRaza.Orco
866                                         CabezaFinal = RandomNumber(717, 719)
    
                                    End Select
    
                            End Select
    
868                         CabezaActual = .OrigChar.Head
                            
870                         .Char.Head = CabezaFinal
872                         .OrigChar.Head = CabezaFinal
874                         Call ChangeUserChar(UserIndex, .Char.Body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    
                            'Quitamos del inv el item
876                         If CabezaActual <> CabezaFinal Then
878                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 102, 0))
880                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
882                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
884                             Call WriteConsoleMsg(UserIndex, "¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
886                     Case 18  ' tan solo crea una particula por determinado tiempo
    
                            Dim Particula           As Integer
    
                            Dim Tiempo              As Long
    
                            Dim ParticulaPermanente As Byte
    
                            Dim sobrechar           As Byte
    
888                         If obj.CreaParticula <> "" Then
890                             Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
892                             Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
894                             ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
896                             sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                                
898                             If ParticulaPermanente = 1 Then
900                                 .Char.ParticulaFx = Particula
902                                 .Char.loops = Tiempo
    
                                End If
                                
904                             If sobrechar = 1 Then
906                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(.Pos.X, .Pos.Y, Particula, Tiempo))
                                Else
                                
908                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, Particula, Tiempo, False))
    
                                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(.Pos.x, .Pos.Y, Particula, Tiempo))
                                End If
    
                            End If
                            
910                         If obj.CreaFX <> 0 Then
912                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, .Pos.X, .Pos.Y))
                                
                                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, obj.CreaFX, 0))
                                ' PrepareMessageCreateFX
                            End If
                            
914                         If obj.Snd1 <> 0 Then
916                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
    
                            End If
                            
918                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
920                     Case 19 ' Reseteo de skill
    
                            Dim S As Byte
                    
922                         If .Stats.UserSkills(eSkill.liderazgo) >= 80 Then
924                             Call WriteConsoleMsg(UserIndex, "Has fundado un clan, no podes resetar tus skills. ", FontTypeNames.FONTTYPE_INFOIAO)
                                Exit Sub
    
                            End If
                        
926                         For S = 1 To NUMSKILLS
928                             .Stats.UserSkills(S) = 0
930                         Next S
                        
                            Dim SkillLibres As Integer
                        
932                         SkillLibres = 5
934                         SkillLibres = SkillLibres + (5 * .Stats.ELV)
                         
936                         .Stats.SkillPts = SkillLibres
938                         Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                        
940                         Call WriteConsoleMsg(UserIndex, "Tus skills han sido reseteados.", FontTypeNames.FONTTYPE_INFOIAO)
942                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
                        ' Mochila
944                     Case 20
                    
946                         If .Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
948                             .Stats.InventLevel = .Stats.InventLevel + 1
950                             .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
952                             Call WriteInventoryUnlockSlots(UserIndex)
954                             Call WriteConsoleMsg(UserIndex, "Has aumentado el espacio de tu inventario!", FontTypeNames.FONTTYPE_INFO)
956                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
958                             Call WriteConsoleMsg(UserIndex, "Ya has desbloqueado todos los casilleros disponibles.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                            
                        ' Poción negra (suicidio)
960                     Case 21
                            'Quitamos del inv el item
962                         Call QuitarUserInvItem(UserIndex, Slot, 1)
                            
964                         If obj.Snd1 <> 0 Then
966                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
968                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                            End If

970                         Call WriteConsoleMsg(UserIndex, "Te has suicidado.", FontTypeNames.FONTTYPE_EJECUCION)
972                         Call UserDie(UserIndex)
                    
                    End Select
    
974                 Call WriteUpdateUserStats(UserIndex)
976                 Call UpdateUserInv(False, UserIndex, Slot)
    
978             Case eOBJType.otBebidas
    
980                 If .flags.Muerto = 1 Then
982                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
984                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
    
986                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
988                 .flags.Sed = 0
990                 Call WriteUpdateHungerAndThirst(UserIndex)
            
                    'Quitamos del inv el item
992                 Call QuitarUserInvItem(UserIndex, Slot, 1)
            
994                 If obj.Snd1 <> 0 Then
996                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                
                    Else
998                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                      End If
            
1000                 Call UpdateUserInv(False, UserIndex, Slot)
            
1002             Case eOBJType.OtCofre
    
1004                 If .flags.Muerto = 1 Then
1006                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
    
                        'Quitamos del inv el item
1008                 Call QuitarUserInvItem(UserIndex, Slot, 1)
1010                 Call UpdateUserInv(False, UserIndex, Slot)
            
1012                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(.Name & " ha abierto un " & obj.Name & " y obtuvo...", FontTypeNames.FONTTYPE_New_DONADOR))
            
1014                 If obj.Snd1 <> 0 Then
1016                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                          End If
            
1018                 If obj.CreaFX <> 0 Then
1020                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, obj.CreaFX, 0))
                          End If
            
                          Dim i As Byte
    
1022                 If obj.Subtipo = 1 Then
    
1024                     For i = 1 To obj.CantItem

1026                        If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                
1028                             If (.flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) Then
1030                                 Call TirarItemAlPiso(.Pos, obj.Item(i))
                                     End If
                                
                                 End If
                            
1032                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).Name & " (" & obj.Item(i).amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))

1034                     Next i
            
                          Else
            
1036                     For i = 1 To obj.CantEntrega
    
                                  Dim indexobj As Byte
1038                            indexobj = RandomNumber(1, obj.CantItem)
                
                                  Dim Index As obj
    
1040                         Index.ObjIndex = obj.Item(indexobj).ObjIndex
1042                         Index.amount = obj.Item(indexobj).amount
    
1044                         If Not MeterItemEnInventario(UserIndex, Index) Then

1046                            If (.flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) Then
1048                                 Call TirarItemAlPiso(.Pos, Index)
                                     End If
                                
                                  End If

1050                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).Name & " (" & Index.amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
1052                     Next i
    
                          End If
        
1054             Case eOBJType.otLlaves
1056                 Call WriteConsoleMsg(UserIndex, "Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.", FontTypeNames.FONTTYPE_INFO)
        
1058             Case eOBJType.otBotellaVacia
    
1060                 If .flags.Muerto = 1 Then
1062                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
    
1064                 If (MapData(.Pos.Map, .flags.TargetX, .flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
1066                     Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
    
1068                 MiObj.amount = 1
1070                 MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta

1072                 Call QuitarUserInvItem(UserIndex, Slot, 1)
    
1074                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
1076                     Call TirarItemAlPiso(.Pos, MiObj)
                          End If
            
1078                 Call UpdateUserInv(False, UserIndex, Slot)
        
1080             Case eOBJType.otBotellaLlena
    
1082                 If .flags.Muerto = 1 Then
1084                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                              ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
    
1086                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
    
1088                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
1090                 .flags.Sed = 0
1092                 Call WriteUpdateHungerAndThirst(UserIndex)
1094                 MiObj.amount = 1
1096                 MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
1098                 Call QuitarUserInvItem(UserIndex, Slot, 1)
    
1100                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
1102                     Call TirarItemAlPiso(.Pos, MiObj)
    
                          End If
            
1104                 Call UpdateUserInv(False, UserIndex, Slot)
        
1106             Case eOBJType.otPergaminos
    
1108                 If .flags.Muerto = 1 Then
1110                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                              ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
                          'Call LogError(.Name & " intento aprender el hechizo " & ObjData(.Invent.Object(slot).ObjIndex).HechizoIndex)
            
1112                 If ClasePuedeUsarItem(UserIndex, .Invent.Object(Slot).ObjIndex, Slot) Then
    
                              'If .Stats.MaxMAN > 0 Then
1114                     If .flags.Hambre = 0 And .flags.Sed = 0 Then
1116                         Call AgregarHechizo(UserIndex, Slot)
1118                         Call UpdateUserInv(False, UserIndex, Slot)
                                  ' Call LogError(.Name & " lo aprendio.")
                              Else
1120                         Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
    
                              End If
    
                              ' Else
                              '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_WARNING)
                              'End If
                          Else
                 
1122                     Call WriteConsoleMsg(UserIndex, "Por mas que lo intentas, no podés comprender el manuescrito.", FontTypeNames.FONTTYPE_INFO)
       
                          End If
            
1124             Case eOBJType.otMinerales
    
1126                 If .flags.Muerto = 1 Then
1128                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
    
1130                 Call WriteWorkRequestTarget(UserIndex, FundirMetal)
           
1132             Case eOBJType.otInstrumentos
    
1134                 If .flags.Muerto = 1 Then
1136                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
1138                 If obj.Real Then '¿Es el Cuerno Real?
1140                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1142                         If MapInfo(.Pos.Map).Seguro = 1 Then
1144                             Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                                      Exit Sub
    
                                  End If
    
1146                         Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                  Exit Sub
                              Else
1148                         Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                                  Exit Sub
    
                              End If
    
1150                 ElseIf obj.Caos Then '¿Es el Cuerno Legión?
    
1152                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1154                         If MapInfo(.Pos.Map).Seguro = 1 Then
1156                             Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                                      Exit Sub
    
                                  End If
    
1158                         Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                  Exit Sub
                              Else
1160                         Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                                  Exit Sub
    
                              End If
    
                          End If
    
                          'Si llega aca es porque es o Laud o Tambor o Flauta
1162                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
           
1164             Case eOBJType.otBarcos
                
                        ' Piratas y trabajadores navegan al nivel 23
                     If .Invent.Object(Slot).ObjIndex <> 199 And .Invent.Object(Slot).ObjIndex <> 200 Then
1166                     If .clase = eClass.Trabajador Or .clase = eClass.Pirat Then
1168                         If .Stats.ELV < 23 Then
1170                             Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 23 o superior.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                End If
                        ' Nivel mínimo 25 para navegar, si no sos pirata ni trabajador
1172                    ElseIf .Stats.ELV < 25 Then
1174                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    Else
                        If MapData(.Pos.Map, .Pos.X + 1, .Pos.Y).trigger <> 8 And MapData(.Pos.Map, .Pos.X - 1, .Pos.Y).trigger <> 8 And MapData(.Pos.Map, .Pos.X, .Pos.Y + 1).trigger <> 8 And MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).trigger <> 8 Then
                            Call WriteConsoleMsg(UserIndex, "Este traje es para aguas contaminadas.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If

1176                If .flags.Navegando = 0 Then
1178                    If LegalWalk(.Pos.Map, .Pos.X - 1, .Pos.Y, eHeading.WEST, True, False) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y - 1, eHeading.NORTH, True, False) Or LegalWalk(.Pos.Map, .Pos.X + 1, .Pos.Y, eHeading.EAST, True, False) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y + 1, eHeading.SOUTH, True, False) Then
1180                        Call DoNavega(UserIndex, obj, Slot)
                             Else
1182                        Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco o traje de baño!", FontTypeNames.FONTTYPE_INFO)
                             End If
                    
                         Else
1184                     If .Invent.BarcoObjIndex <> .Invent.Object(Slot).ObjIndex Then
1186                        Call DoNavega(UserIndex, obj, Slot)
                             Else
1188                        If LegalWalk(.Pos.Map, .Pos.X - 1, .Pos.Y, eHeading.WEST, False, True) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y - 1, eHeading.NORTH, False, True) Or LegalWalk(.Pos.Map, .Pos.X + 1, .Pos.Y, eHeading.EAST, False, True) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y + 1, eHeading.SOUTH, False, True) Then
1190                            Call DoNavega(UserIndex, obj, Slot)
                                 Else
1192                            Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte a la costa para dejar la barca!", FontTypeNames.FONTTYPE_INFO)
                                 End If
                             End If
                         End If
            
1194             Case eOBJType.otMonturas
                          'Verifica todo lo que requiere la montura
        
1196                If .flags.Muerto = 1 Then
1198                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡Estas muerto! Los fantasmas no pueden montar.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
                
1200                If .flags.Navegando = 1 Then
1202                    Call WriteConsoleMsg(UserIndex, "Debes dejar de navegar para poder cabalgar.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1204                If MapInfo(.Pos.Map).zone = "DUNGEON" Then
1206                    Call WriteConsoleMsg(UserIndex, "No podes cabalgar dentro de un dungeon.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1208                Call DoMontar(UserIndex, obj, Slot)
    
1210             Case eOBJType.OtDonador
    
1212                 Select Case obj.Subtipo
    
                              Case 1
                
1214                         If .Counters.Pena <> 0 Then
1216                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                                      Exit Sub
    
                                  End If
                    
1218                         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
1220                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                                      Exit Sub
    
                                  End If
                
1222                         Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1224                         Call WriteConsoleMsg(UserIndex, "Has viajado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
1226                         Call QuitarUserInvItem(UserIndex, Slot, 1)
1228                         Call UpdateUserInv(False, UserIndex, Slot)
                    
1230                     Case 2
    
1232                         If DonadorCheck(.Cuenta) = 0 Then
1234                             Call DonadorTiempo(.Cuenta, CLng(obj.CuantoAumento))
1236                             Call WriteConsoleMsg(UserIndex, "Donación> Se han agregado " & obj.CuantoAumento & " dias de donador a tu cuenta. Relogea tu personaje para empezar a disfrutar la experiencia.", FontTypeNames.FONTTYPE_WARNING)
1238                             Call QuitarUserInvItem(UserIndex, Slot, 1)
1240                             Call UpdateUserInv(False, UserIndex, Slot)
                                  Else
1242                             Call DonadorTiempo(.Cuenta, CLng(obj.CuantoAumento))
1244                             Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & CLng(obj.CuantoAumento) & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
1246                             .donador.activo = 1
1248                             Call QuitarUserInvItem(UserIndex, Slot, 1)
1250                             Call UpdateUserInv(False, UserIndex, Slot)
    
                                      'Call WriteConsoleMsg(UserIndex, "Donación> Debes esperar a que finalice el periodo existente para renovar tu suscripción.", FontTypeNames.FONTTYPE_INFOIAO)
                                  End If
    
1252                     Case 3
1254                         Call AgregarCreditosDonador(.Cuenta, CLng(obj.CuantoAumento))
1256                         Call WriteConsoleMsg(UserIndex, "Donación> Tu credito ahora es de " & CreditosDonadorCheck(.Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
1258                         Call QuitarUserInvItem(UserIndex, Slot, 1)
1260                         Call UpdateUserInv(False, UserIndex, Slot)
    
                          End Select
         
1262             Case eOBJType.otpasajes
    
1264                 If .flags.Muerto = 1 Then
1266                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
1268                 If .flags.TargetNpcTipo <> Pirata Then
1270                     Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
1272                 If Distancia(NpcList(.flags.TargetNPC).Pos, .Pos) > 3 Then
1274                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
1276                 If .Pos.Map <> obj.DesdeMap Then
                              Rem  Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aquí! Largate!", FontTypeNames.FONTTYPE_INFO)
1278                     Call WriteChatOverHead(UserIndex, "El pasaje no lo compraste aquí! Largate!", str(NpcList(.flags.TargetNPC).Char.CharIndex), vbWhite)
                              Exit Sub
    
                          End If
            
1280                 If Not MapaValido(obj.HastaMap) Then
                              Rem Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
1282                     Call WriteChatOverHead(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str(NpcList(.flags.TargetNPC).Char.CharIndex), vbWhite)
                              Exit Sub
    
                          End If
    
1284                 If obj.NecesitaNave > 0 Then
1286                     If .Stats.UserSkills(eSkill.Navegacion) < 80 Then
                                  Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
1288                         Call WriteChatOverHead(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str(NpcList(.flags.TargetNPC).Char.CharIndex), vbWhite)
                                  Exit Sub
    
                              End If
    
                          End If
                
1290                 Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1292                 Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
1294                 .Stats.MinAGU = 0
1296                 .Stats.MinHam = 0
1298                 .flags.Sed = 1
1300                 .flags.Hambre = 1
1302                 Call WriteUpdateHungerAndThirst(UserIndex)
1304                 Call QuitarUserInvItem(UserIndex, Slot, 1)
1306                 Call UpdateUserInv(False, UserIndex, Slot)
            
1308             Case eOBJType.otRunas
        
1310                If .Counters.Pena <> 0 Then
1312                    Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1314                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
1316                    Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
                        
1318                If MapInfo(.Pos.Map).Seguro = 0 And .flags.Muerto = 0 Then
1320                    Call WriteConsoleMsg(UserIndex, "Solo podes usar tu runa en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1322                If .Accion.AccionPendiente Then
                           Exit Sub
    
                       End If
            
1324                 Select Case ObjData(ObjIndex).TipoRuna
            
                              Case 1, 2
    
1326                         If Not EsGM(UserIndex) Then
1328                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Runa, 400, False))
1330                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 350, Accion_Barra.Runa))
                                  Else
1332                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Runa, 50, False))
1334                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 100, Accion_Barra.Runa))
    
                                  End If
    
1336                         .Accion.Particula = ParticulasIndex.Runa
1338                         .Accion.AccionPendiente = True
1340                         .Accion.TipoAccion = Accion_Barra.Runa
1342                         .Accion.RunaObj = ObjIndex
1344                         .Accion.ObjSlot = Slot
        
                          End Select
            
1346             Case eOBJType.otmapa
1348                 Call WriteShowFrmMapa(UserIndex)
            
                  End Select
             
             End With

             Exit Sub

hErr:
1350    LogError "Error en useinvitem Usuario: " & UserList(UserIndex).Name & " item:" & obj.Name & " index: " & UserList(UserIndex).Invent.Object(Slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmasConstruibles_Err
        

100     Call WriteBlacksmithWeapons(UserIndex)

        
        Exit Sub

EnivarArmasConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmasConstruibles", Erl)

        
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruibles_Err
        

100     Call WriteCarpenterObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruibles", Erl)

        
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesAlquimia_Err
        

100     Call WriteAlquimistaObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesAlquimia_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)

        
End Sub

Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesSastre_Err
        

100     Call WriteSastreObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesSastre_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)

        
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmadurasConstruibles_Err
        

100     Call WriteBlacksmithArmors(UserIndex)

        
        Exit Sub

EnivarArmadurasConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmadurasConstruibles", Erl)

        
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemSeCae_Err
        

100     ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> eOBJType.otLlaves And ObjData(Index).OBJType <> eOBJType.otBarcos And ObjData(Index).OBJType <> eOBJType.otMonturas And ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And ObjData(Index).donador = 0 And Not ObjData(Index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemSeCae", Erl)

        
End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo PirataCaeItem_Err

100     With UserList(UserIndex)

102       If .clase = eClass.Pirat And .Stats.ELV >= 37 And .flags.Navegando = 1 Then

            ' Si no está navegando, se caen los items
104         If .Invent.BarcoObjIndex > 0 Then

                ' Con galeón cada item tiene una probabilidad de caerse del 67%
106             If ObjData(.Invent.BarcoObjIndex).Ropaje = iGaleon Then

108                 If RandomNumber(1, 100) <= 33 Then
                        Exit Function
                      End If

                    End If

                End If

            End If

        End With

110     PirataCaeItem = True

        Exit Function

PirataCaeItem_Err:
112     Call TraceError(Err.Number, Err.Description, "InvUsuario.PirataCaeItem", Erl)


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
        
        On Error GoTo TirarTodosLosItems_Err

        Dim i         As Byte
        Dim NuevaPos  As WorldPos
        Dim MiObj     As obj
        Dim ItemIndex As Integer
       
100     With UserList(UserIndex)
            ' Tambien se cae el oro de la billetera
102         If (.Stats.GLD < 100000) Then
104             Call TirarOro(.Stats.GLD, UserIndex)
            End If
            
106         For i = 1 To .CurrentInventorySlots
    
108             ItemIndex = .Invent.Object(i).ObjIndex

110             If ItemIndex > 0 Then

112                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
114                     NuevaPos.X = 0
116                     NuevaPos.Y = 0
                    
118                     MiObj.amount = .Invent.Object(i).amount
120                     MiObj.ObjIndex = ItemIndex

122                     If .flags.CarroMineria = 1 Then
124                         If ItemIndex = ORO_MINA Or ItemIndex = PLATA_MINA Or ItemIndex = HIERRO_MINA Then
126                             MiObj.amount = Int(MiObj.amount * 0.3)
                            End If
                        End If
                    
128                     Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
            
130                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
132                         Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        
                        ' WyroX: Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                        Else
134                         Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
136                         Call UpdateUserInv(False, UserIndex, i)
                        End If
                
                    End If

                End If
    
138         Next i
    
        End With
 
        Exit Sub

TirarTodosLosItems_Err:
140     Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarTodosLosItems", Erl)


        
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo ItemNewbie_Err
        

100     ItemNewbie = ObjData(ItemIndex).Newbie = 1

        
        Exit Function

ItemNewbie_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemNewbie", Erl)

        
End Function
