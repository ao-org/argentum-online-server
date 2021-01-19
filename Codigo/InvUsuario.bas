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
    
        

        '17/09/02
        'Agregue que la función se asegure que el objeto no es un barco

        

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
112     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.TieneObjetosRobables", Erl)

        
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional slot As Byte) As Boolean

        On Error GoTo manejador

        'Call LogTarea("ClasePuedeUsarItem")

        Dim flag As Boolean

100     If slot <> 0 Then
102         If UserList(UserIndex).Invent.Object(slot).Equipped Then
104             ClasePuedeUsarItem = True
                Exit Function

            End If

        End If

106     If EsGM(UserIndex) Then
108         ClasePuedeUsarItem = True
            Exit Function

        End If

        'Admins can use ANYTHING!
        'If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        'If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer

110     For i = 1 To NUMCLASES

112         If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
114             ClasePuedeUsarItem = False
                Exit Function

            End If

116     Next i

        ' End If
        'End If

118     ClasePuedeUsarItem = True

        Exit Function

manejador:
120     LogError ("Error en ClasePuedeUsarItem")

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
            
126             Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
128                 DeDonde = Lindos
                
130             Case eCiudad.cArghal 'Vamos a tener que ir por todo el desierto... uff!
132                 DeDonde = Arghal
                
134             Case eCiudad.CHillidan
136                 DeDonde = Hillidan
                
138             Case Else
140                 DeDonde = Ullathorpe

            End Select
        
142         Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
        End If

        
        Exit Sub

QuitarNewbieObj_Err:
144     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.QuitarNewbieObj", Erl)
146     Resume Next
        
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
        
        On Error GoTo LimpiarInventario_Err
        

        Dim j As Integer

100     For j = 1 To UserList(UserIndex).CurrentInventorySlots
102         UserList(UserIndex).Invent.Object(j).ObjIndex = 0
104         UserList(UserIndex).Invent.Object(j).Amount = 0
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
158     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.LimpiarInventario", Erl)
160     Resume Next
        
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
104             Call LogGM(.name, " trató de tirar " & PonerPuntos(Cantidad) & " de oro en " & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
                Exit Sub

            End If
        
            'If Cantidad > 100000 Then Exit Sub
106         If .flags.BattleModo = 1 Then Exit Sub
        
            'SI EL Pjta TIENE ORO LO TIRAMOS
108         If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then

                Dim i     As Byte
                Dim MiObj As obj

                'info debug
                Dim loops As Integer
                Dim Extra    As Long
                Dim TeniaOro As Long

110             TeniaOro = .Stats.GLD

112             If Cantidad > 500000 Then 'Para evitar explotar demasiado
114                 Extra = Cantidad - 500000
116                 Cantidad = 500000

                End If
        
118             Do While (Cantidad > 0)
            
120                 If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
122                     MiObj.Amount = MAX_INVENTORY_OBJS
124                     Cantidad = Cantidad - MiObj.Amount
                    Else
126                     MiObj.Amount = Cantidad
128                     Cantidad = Cantidad - MiObj.Amount

                    End If

130                 MiObj.ObjIndex = iORO

                    Dim AuxPos As WorldPos

132                 If .clase = eClass.Pirat Then
134                     AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    Else
136                     AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    End If
            
138                 If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
140                     .Stats.GLD = .Stats.GLD - MiObj.Amount

                    End If
            
                    'info debug
142                 loops = loops + 1

144                 If loops > 100 Then
146                     Call LogError("Se ha superado el limite de iteraciones(100) permitido en el Sub TirarOro()")
                        Exit Sub

                    End If
            
                Loop
                
                ' Si es GM, registramos lo q hizo incluso si es Horacio
148             If EsGM(UserIndex) Then

150                 If MiObj.ObjIndex = iORO Then
152                     Call LogGM(.name, "Tiro: " & PonerPuntos(Cantidad) & " monedas de oro.")

                    Else
154                     Call LogGM(.name, "Tiro cantidad:" & PonerPuntos(Cantidad) & " Objeto:" & ObjData(MiObj.ObjIndex).name)

                    End If

                End If
        
156             If TeniaOro = .Stats.GLD Then Extra = 0

158             If Extra > 0 Then
160                 .Stats.GLD = .Stats.GLD - Extra
                End If
    
            End If
        
        End With

        Exit Sub

ErrHandler:
162 Call RegistrarError(Err.Number, Err.Description, "InvUsuario.TirarOro", Erl())
    
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarUserInvItem_Err
        

100     If slot < 1 Or slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
102     With UserList(UserIndex).Invent.Object(slot)

104         If .Amount <= Cantidad And .Equipped = 1 Then
106             Call Desequipar(UserIndex, slot)

            End If
        
            'Quita un objeto
108         .Amount = .Amount - Cantidad

            '¿Quedan mas?
110         If .Amount <= 0 Then
112             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
114             .ObjIndex = 0
116             .Amount = 0

            End If

        End With

        
        Exit Sub

QuitarUserInvItem_Err:
118     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.QuitarUserInvItem", Erl)
120     Resume Next
        
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateUserInv_Err
        

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(UserIndex).Invent.Object(slot).ObjIndex > 0 Then
104             Call ChangeUserInv(UserIndex, slot, UserList(UserIndex).Invent.Object(slot))
            Else
106             Call ChangeUserInv(UserIndex, slot, NullObj)

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
118     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.UpdateUserInv", Erl)
120     Resume Next
        
End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        
        On Error GoTo DropObj_Err

        Dim obj As obj

100     If num > 0 Then
            
102         With UserList(UserIndex)

104             If num > .Invent.Object(slot).Amount Then
106                 num = .Invent.Object(slot).Amount
                End If
    
108             obj.ObjIndex = .Invent.Object(slot).ObjIndex
110             obj.Amount = num
    
112             If ObjData(obj.ObjIndex).Destruye = 0 Then
    
                    'Check objeto en el suelo
114                 If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Then
                      
116                     If num + MapData(.Pos.Map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
118                         num = MAX_INVENTORY_OBJS - MapData(.Pos.Map, X, Y).ObjInfo.Amount
                        End If
                        
                        ' Si sos Admin, Dios o Usuario, crea el objeto en el piso.
120                     If (.flags.Privilegios And (PlayerType.user Or PlayerType.Admin Or PlayerType.Dios)) <> 0 Then

                            ' Tiramos el item al piso
122                         Call MakeObj(obj, Map, X, Y)

                        End If
                        
124                     Call QuitarUserInvItem(UserIndex, slot, num)
126                     Call UpdateUserInv(False, UserIndex, slot)
                        
128                     If Not .flags.Privilegios And PlayerType.user Then
130                         Call LogGM(.name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)
                        End If
    
                    Else
                    
                        'Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
132                     Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                Else
134                 Call QuitarUserInvItem(UserIndex, slot, num)
136                 Call UpdateUserInv(False, UserIndex, slot)
    
                End If
            
            End With

        End If
        
        Exit Sub

DropObj_Err:
138     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.DropObj", Erl)

140     Resume Next
        
End Sub

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo EraseObj_Err
        

        Dim Rango As Byte

100     MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount - num

102     If MapData(Map, X, Y).ObjInfo.Amount <= 0 Then

            'Rango = val(ReadField(1, ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).CreaLuz, Asc(":")))
    
            ' If Rango >= 1 Then
            '  'Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageLightFXToFloor(X, Y, 0, Rango))
            '  MapData(Map, x, Y).Luz.Color = 0
            '   MapData(Map, x, Y).Luz.Rango = 0
            ' End If
    
            '  If ObjData(MapData(Map, x, Y).ObjInfo.ObjIndex).CreaParticulaPiso >= 1 Then
            ' Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageParticleFXToFloor(X, Y, 0, 0))
            '   MapData(Map, x, Y).Particula = 0
            '   MapData(Map, x, Y).TimeParticula = 0
            ' End If

            

104         If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType <> otTeleport Then
106             Call QuitarItemLimpieza(Map, X, Y)
            End If
            
108         MapData(Map, X, Y).ObjInfo.ObjIndex = 0
110         MapData(Map, X, Y).ObjInfo.Amount = 0
    
    
112         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

        
        Exit Sub

EraseObj_Err:
114     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.EraseObj", Erl)
116     Resume Next
        
End Sub

Sub MakeObj(ByRef obj As obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Limpiar As Boolean = True)
        
        On Error GoTo MakeObj_Err

        Dim Color As Long

        Dim Rango As Byte

100     If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
102         If MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex Then
104             MapData(Map, X, Y).ObjInfo.Amount = MapData(Map, X, Y).ObjInfo.Amount + obj.Amount
            Else
                ' Lo agrego a la limpieza del mundo o reseteo el timer si el objeto ya existía
106             If ObjData(obj.ObjIndex).OBJType <> otTeleport Then
108                 Call AgregarItemLimpieza(Map, X, Y, MapData(Map, X, Y).ObjInfo.ObjIndex <> 0)
                End If
            
110             MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex

112             If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
114                 MapData(Map, X, Y).ObjInfo.Amount = ObjData(obj.ObjIndex).VidaUtil
                Else
116                 MapData(Map, X, Y).ObjInfo.Amount = obj.Amount

                End If
            
                'Color = val(ReadField(2, ObjData(obj.ObjIndex).CreaLuz, Asc(":")))
                ' Rango = val(ReadField(1, ObjData(obj.ObjIndex).CreaLuz, Asc(":")))
    
                ' If Rango >= 1 Then
                'Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageLightFXToFloor(X, Y, color, Rango))
                '  MapData(Map, x, Y).Luz.Color = Color
                '  MapData(Map, x, Y).Luz.Rango = Rango
                ' End If
    
                ' If ObjData(obj.ObjIndex).CreaParticulaPiso >= 1 Then
                'Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageParticleFXToFloor(X, Y, ObjData(obj.ObjIndex).CreaParticulaPiso, -1))
                ' MapData(Map, x, Y).Particula = ObjData(obj.ObjIndex).CreaParticulaPiso
                ' MapData(Map, x, Y).TimeParticula = -1
                ' End If
118             Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(obj.ObjIndex, X, Y))
                
            End If
    
        End If
        
        Exit Sub

MakeObj_Err:
120     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.MakeObj", Erl)

122     Resume Next
        
End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean

        On Error GoTo ErrHandler

        'Call LogTarea("MeterItemEnInventario")
 
        Dim X    As Integer

        Dim Y    As Integer

        Dim slot As Byte

        '¿el user ya tiene un objeto del mismo tipo? ?????
100     If MiObj.ObjIndex = 12 Then
102         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount
            Call WriteUpdateGold(UserIndex)

        Else
    
104         slot = 1

106         Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
108             slot = slot + 1

110             If slot > UserList(UserIndex).CurrentInventorySlots Then
                    Exit Do

                End If

            Loop
        
            'Sino busca un slot vacio
112         If slot > UserList(UserIndex).CurrentInventorySlots Then
114             slot = 1

116             Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
118                 slot = slot + 1

120                 If slot > UserList(UserIndex).CurrentInventorySlots Then
                        'Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
122                     Call WriteLocaleMsg(UserIndex, "328", FontTypeNames.FONTTYPE_FIGHT)
124                     MeterItemEnInventario = False
                        Exit Function

                    End If

                Loop
126             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

            End If
        
            'Mete el objeto
128         If UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
                'Menor que MAX_INV_OBJS
130             UserList(UserIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex
132             UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount
            Else
134             UserList(UserIndex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS

            End If
        
136         MeterItemEnInventario = True
           
138         Call UpdateUserInv(False, UserIndex, slot)

        End If

142     MeterItemEnInventario = True

        Exit Function
ErrHandler:

End Function

Function MeterItemEnInventarioDeNpc(ByVal NpcIndex As Integer, ByRef MiObj As obj) As Boolean

        On Error GoTo ErrHandler

        'Call LogTarea("MeterItemEnInventario")
 
        Dim X    As Integer

        Dim Y    As Integer

        Dim slot As Byte

        '¿el user ya tiene un objeto del mismo tipo? ?????
    
100     slot = 1

102     Do Until Npclist(NpcIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And Npclist(NpcIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
104         slot = slot + 1

106         If slot > MAX_INVENTORY_SLOTS Then
                Exit Do

            End If

        Loop
        
        'Sino busca un slot vacio
108     If slot > MAX_INVENTORY_SLOTS Then
110         slot = 1

112         Do Until Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
114             slot = slot + 1

116             If slot > MAX_INVENTORY_SLOTS Then
                    Rem Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
118                 MeterItemEnInventarioDeNpc = False
                    Exit Function

                End If

            Loop
120         Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

        End If

122     MeterItemEnInventarioDeNpc = True

        Exit Function
ErrHandler:

End Function

Sub GetObj(ByVal UserIndex As Integer)
        
        On Error GoTo GetObj_Err
        

        Dim obj   As ObjData

        Dim MiObj As obj

        '¿Hay algun obj?
100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
102         If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

                Dim X    As Integer

                Dim Y    As Integer

                Dim slot As Byte
        
104             X = UserList(UserIndex).Pos.X
106             Y = UserList(UserIndex).Pos.Y
108             obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
110             MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount
112             MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex
        
114             If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Else
            
                    'Quitamos el objeto
116                 Call EraseObj(MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

118                 If Not UserList(UserIndex).flags.Privilegios And PlayerType.user Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
    
120                 If BusquedaTesoroActiva Then
122                     If UserList(UserIndex).Pos.Map = TesoroNumMapa And UserList(UserIndex).Pos.X = TesoroX And UserList(UserIndex).Pos.Y = TesoroY Then
    
124                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).name & " encontro el tesoro ¡Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
126                         BusquedaTesoroActiva = False

                        End If

                    End If
                
128                 If BusquedaRegaloActiva Then
130                     If UserList(UserIndex).Pos.Map = RegaloNumMapa And UserList(UserIndex).Pos.X = RegaloX And UserList(UserIndex).Pos.Y = RegaloY Then
132                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).name & " fue el valiente que encontro el gran item magico ¡Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
134                         BusquedaRegaloActiva = False

                        End If

                    End If
                
                    'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                    'Es un Objeto que tenemos que loguear?
136                 If ObjData(MiObj.ObjIndex).Log = 1 Then
138                     Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)

                        ' ElseIf MiObj.Amount = 1000 Then 'Es mucha cantidad?
                        '  'Si no es de los prohibidos de loguear, lo logueamos.
                        '   'If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                        ' Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
                        ' End If
                    End If
                
                End If

            End If

        Else

140         If Not UserList(UserIndex).flags.UltimoMensaje = 261 Then
142             Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)
144             UserList(UserIndex).flags.UltimoMensaje = 261

            End If
    
            'Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
        End If

        
        Exit Sub

GetObj_Err:
146     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.GetObj", Erl)
148     Resume Next
        
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo Desequipar_Err

        'Desequipa el item slot del inventario
        Dim obj As ObjData

100     If (slot < LBound(UserList(UserIndex).Invent.Object)) Or (slot > UBound(UserList(UserIndex).Invent.Object)) Then
            Exit Sub
102     ElseIf UserList(UserIndex).Invent.Object(slot).ObjIndex = 0 Then
            Exit Sub

        End If

104     obj = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex)

106     Select Case obj.OBJType

            Case eOBJType.otWeapon
108             UserList(UserIndex).Invent.Object(slot).Equipped = 0
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
130             UserList(UserIndex).Invent.Object(slot).Equipped = 0
132             UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
134             UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
                ' Case eOBJType.otAnillos
                '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
                '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
                ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
            
136         Case eOBJType.otHerramientas
138             UserList(UserIndex).Invent.Object(slot).Equipped = 0
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
162                     UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                
164                     UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                        ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
166                     Call WriteFYA(UserIndex)

168                 Case 3 'Modifica los skills
170                     UserList(UserIndex).Stats.UserSkills(obj.QueSkill) = UserList(UserIndex).Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento

172                 Case 4 ' Regeneracion Vida
174                     UserList(UserIndex).flags.RegeneracionHP = 0

176                 Case 5 ' Regeneracion Mana
178                     UserList(UserIndex).flags.RegeneracionMana = 0

180                 Case 6 'Aumento Golpe
182                     UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit - obj.CuantoAumento
184                     UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - obj.CuantoAumento

186                 Case 7 '
                
188                 Case 9 ' Orbe Ignea
190                     UserList(UserIndex).flags.NoMagiaEfeceto = 0

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

222                 Case 19
224                     UserList(UserIndex).flags.Envenena = 0

226                 Case 20 ' anillo de las sombras
228                     UserList(UserIndex).flags.AnilloOcultismo = 0
                
                End Select
        
230             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 5))
232             UserList(UserIndex).Char.Otra_Aura = 0
234             UserList(UserIndex).Invent.Object(slot).Equipped = 0
236             UserList(UserIndex).Invent.MagicoObjIndex = 0
238             UserList(UserIndex).Invent.MagicoSlot = 0
        
240         Case eOBJType.otNUDILLOS
    
                'falta mandar animacion
            
242             UserList(UserIndex).Invent.Object(slot).Equipped = 0
244             UserList(UserIndex).Invent.NudilloObjIndex = 0
246             UserList(UserIndex).Invent.NudilloSlot = 0
        
248             UserList(UserIndex).Char.Arma_Aura = ""
250             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 1))
        
252             UserList(UserIndex).Char.WeaponAnim = NingunArma
254             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        
256         Case eOBJType.otArmadura
258             UserList(UserIndex).Invent.Object(slot).Equipped = 0
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
    
280         Case eOBJType.otCASCO
282             UserList(UserIndex).Invent.Object(slot).Equipped = 0
284             UserList(UserIndex).Invent.CascoEqpObjIndex = 0
286             UserList(UserIndex).Invent.CascoEqpSlot = 0
288             UserList(UserIndex).Char.Head_Aura = 0
290             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 4))

292             UserList(UserIndex).Char.CascoAnim = NingunCasco
294             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    
296             If obj.ResistenciaMagica > 0 Then
298                 Call WriteUpdateRM(UserIndex)
                End If
    
300         Case eOBJType.otESCUDO
302             UserList(UserIndex).Invent.Object(slot).Equipped = 0
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
324             UserList(UserIndex).Invent.Object(slot).Equipped = 0
326             UserList(UserIndex).Invent.DañoMagicoEqpObjIndex = 0
328             UserList(UserIndex).Invent.DañoMagicoEqpSlot = 0
330             UserList(UserIndex).Char.DM_Aura = 0
332             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 6))
334             Call WriteUpdateDM(UserIndex)
                
336         Case eOBJType.otResistencia
338             UserList(UserIndex).Invent.Object(slot).Equipped = 0
340             UserList(UserIndex).Invent.ResistenciaEqpObjIndex = 0
342             UserList(UserIndex).Invent.ResistenciaEqpSlot = 0
344             UserList(UserIndex).Char.RM_Aura = 0
346             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 7))
348             Call WriteUpdateRM(UserIndex)
        
        End Select

350     Call UpdateUserInv(False, UserIndex, slot)

        
        Exit Sub

Desequipar_Err:
352     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.Desequipar", Erl)
354     Resume Next
        
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
        

104     If ObjData(ObjIndex).Real = 1 Then
106         If Status(UserIndex) = 3 Then
108             FaccionPuedeUsarItem = esArmada(UserIndex)
            Else
110             FaccionPuedeUsarItem = False

            End If

112     ElseIf ObjData(ObjIndex).Caos = 1 Then

114         If Status(UserIndex) = 2 Then
116             FaccionPuedeUsarItem = esCaos(UserIndex)
            Else
118             FaccionPuedeUsarItem = False

            End If

        Else
120         FaccionPuedeUsarItem = True

        End If

        
        Exit Function

FaccionPuedeUsarItem_Err:
122     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.FaccionPuedeUsarItem", Erl)
124     Resume Next
        
End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)

        On Error GoTo ErrHandler

        Dim errordesc As String

        'Equipa un item del inventario
        Dim obj       As ObjData
        Dim ObjIndex  As Integer

100     ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
102     obj = ObjData(ObjIndex)

104     If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
106         Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

108     If UserList(UserIndex).Stats.ELV < obj.MinELV And Not EsGM(UserIndex) Then
110         Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
112     If obj.SkillIndex > 0 Then
    
114         If UserList(UserIndex).Stats.UserSkills(obj.SkillIndex) < obj.SkillRequerido And Not EsGM(UserIndex) Then
116             Call WriteConsoleMsg(UserIndex, "Necesitas " & obj.SkillRequerido & " puntos en " & SkillsNames(obj.SkillIndex) & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

        End If
    
118     With UserList(UserIndex)
    
120         Select Case obj.OBJType

                Case eOBJType.otWeapon
                
122                 errordesc = "Arma"

124                 If Not ClasePuedeUsarItem(UserIndex, ObjIndex, slot) And FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
126                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
128                 If Not FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
130                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Si esta equipado lo quita
132                 If .Invent.Object(slot).Equipped Then
                    
                        'Quitamos del inv el item
134                     Call Desequipar(UserIndex, slot)
                        
                        'Animacion por defecto
136                     .Char.WeaponAnim = NingunArma

138                     If .flags.Montado = 0 Then
140                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If

                        Exit Sub

                    End If
            
                    'Quitamos el elemento anterior
142                 If .Invent.WeaponEqpObjIndex > 0 Then
144                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
            
146                 If .Invent.HerramientaEqpObjIndex > 0 Then
148                     Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
                    End If
            
150                 If .Invent.NudilloObjIndex > 0 Then
152                     Call Desequipar(UserIndex, .Invent.NudilloSlot)
                    End If
            
154                 .Invent.Object(slot).Equipped = 1
156                 .Invent.WeaponEqpObjIndex = .Invent.Object(slot).ObjIndex
158                 .Invent.WeaponEqpSlot = slot
            
160                 If obj.Proyectil = 1 Then 'Si es un arco, desequipa el escudo.
            
                        'If .Invent.EscudoEqpObjIndex = 404 Or .Invent.EscudoEqpObjIndex = 1007 Or .Invent.EscudoEqpObjIndex = 1358 Then
162                     If .Invent.EscudoEqpObjIndex = 1700 Or _
                           .Invent.EscudoEqpObjIndex = 1730 Or _
                           .Invent.EscudoEqpObjIndex = 1724 Or _
                           .Invent.EscudoEqpObjIndex = 1717 Or _
                           .Invent.EscudoEqpObjIndex = 1699 Then
                
                        Else

164                         If .Invent.EscudoEqpObjIndex > 0 Then
166                             Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
168                             Call WriteConsoleMsg(UserIndex, "No podes tirar flechas si tenés un escudo equipado. Tu escudo fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        End If

                    End If
            
                    'Sonido
170                 If obj.SndAura = 0 Then
172                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
174                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
            
176                 If Len(obj.CreaGRH) <> 0 Then
178                     .Char.Arma_Aura = obj.CreaGRH
180                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If
                
182                 If obj.MagicDamageBonus > 0 Then
184                     Call WriteUpdateDM(UserIndex)
                    End If
                
186                 If .flags.Montado = 0 Then
                
188                     If .flags.Navegando = 0 Then
190                         .Char.WeaponAnim = obj.WeaponAnim
192                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
      
194             Case eOBJType.otHerramientas
        
196                 If Not ClasePuedeUsarItem(UserIndex, ObjIndex, slot) Then
198                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
200                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
202                     Call Desequipar(UserIndex, slot)
                        Exit Sub

                    End If

204                 If obj.MinSkill <> 0 Then
                
206                     If .Stats.UserSkills(obj.QueSkill) < obj.MinSkill Then
208                         Call WriteConsoleMsg(UserIndex, "Para podes usar " & obj.name & " necesitas al menos " & obj.MinSkill & " puntos en " & SkillsNames(obj.QueSkill) & ".", FontTypeNames.FONTTYPE_INFOIAO)
                            Exit Sub
                        End If

                    End If

                    'Quitamos el elemento anterior
210                 If .Invent.HerramientaEqpObjIndex > 0 Then
212                     Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
                    End If
             
214                 If .Invent.WeaponEqpObjIndex > 0 Then
216                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
             
218                 .Invent.Object(slot).Equipped = 1
220                 .Invent.HerramientaEqpObjIndex = ObjIndex
222                 .Invent.HerramientaEqpSlot = slot
             
224                 If .flags.Montado = 0 Then
                
226                     If .flags.Navegando = 0 Then
228                         .Char.WeaponAnim = obj.WeaponAnim
230                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
       
232             Case eOBJType.otMagicos
            
234                 errordesc = "Magico"
    
236                 If .flags.Muerto = 1 Then
238                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
                    'Si esta equipado lo quita
240                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
242                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
244                 If .Invent.MagicoObjIndex > 0 Then
246                     Call Desequipar(UserIndex, .Invent.MagicoSlot)
                    End If
        
248                 .Invent.Object(slot).Equipped = 1
250                 .Invent.MagicoObjIndex = .Invent.Object(slot).ObjIndex
252                 .Invent.MagicoSlot = slot
                
                    ' Debug.Print "magico" & obj.EfectoMagico
254                 Select Case obj.EfectoMagico

                        Case 1 ' Regenera Stamina
256                         .flags.RegeneracionSta = 1

258                     Case 2 'Modif la fuerza, agilidad, carisma, etc
                            ' .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo)
260                         .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
                        
262                         .Stats.UserAtributos(obj.QueAtributo) = MinimoInt(.Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento, .Stats.UserAtributosBackUP(obj.QueAtributo) * 2)
                
264                         Call WriteFYA(UserIndex)

266                     Case 3 'Modifica los skills
            
268                         .Stats.UserSkills(obj.QueSkill) = .Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento

270                     Case 4
272                         .flags.RegeneracionHP = 1

274                     Case 5
276                         .flags.RegeneracionMana = 1

278                     Case 6
                            'Call WriteConsoleMsg(UserIndex, "Item, temporalmente deshabilitado.", FontTypeNames.FONTTYPE_INFO)
280                         .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
282                         .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento

284                     Case 9
286                         .flags.NoMagiaEfeceto = 1

288                     Case 10
290                         .flags.incinera = 1

292                     Case 11
294                         .flags.Paraliza = 1

296                     Case 12
298                         .flags.CarroMineria = 1
                
300                     Case 14
                            '.flags.DañoMagico = obj.CuantoAumento
                
302                     Case 15 'Pendiete del Sacrificio
304                         .flags.PendienteDelSacrificio = 1

306                     Case 16
308                         .flags.NoPalabrasMagicas = 1

310                     Case 17
312                         .flags.NoDetectable = 1
                   
314                     Case 18 ' Pendiente del Experto
316                         .flags.PendienteDelExperto = 1

318                     Case 19
320                         .flags.Envenena = 1

322                     Case 20 'Anillo ocultismo
324                         .flags.AnilloOcultismo = 1
    
                    End Select
            
                    'Sonido
326                 If obj.SndAura <> 0 Then
328                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
            
330                 If Len(obj.CreaGRH) <> 0 Then
332                     .Char.Otra_Aura = obj.CreaGRH
334                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
                    End If
        
                    'Call WriteUpdateExp(UserIndex)
                    'Call CheckUserLevel(UserIndex)
            
336             Case eOBJType.otNUDILLOS
    
338                 If .flags.Muerto = 1 Then
340                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
342                 If Not ClasePuedeUsarItem(UserIndex, ObjIndex, slot) Then
344                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                 
346                 If .Invent.WeaponEqpObjIndex > 0 Then
348                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                    End If

350                 If .Invent.Object(slot).Equipped Then
352                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
354                 If .Invent.NudilloObjIndex > 0 Then
356                     Call Desequipar(UserIndex, .Invent.NudilloSlot)

                    End If
        
358                 .Invent.Object(slot).Equipped = 1
360                 .Invent.NudilloObjIndex = .Invent.Object(slot).ObjIndex
362                 .Invent.NudilloSlot = slot
        
                    'Falta enviar anim
364                 If .flags.Montado = 0 Then
                
366                     If .flags.Navegando = 0 Then
368                         .Char.WeaponAnim = obj.WeaponAnim
370                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
            
372                 If obj.SndAura = 0 Then
374                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
376                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
                 
378                 If Len(obj.CreaGRH) <> 0 Then
380                     .Char.Arma_Aura = obj.CreaGRH
382                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If
    
384             Case eOBJType.otFlechas

386                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Or Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
388                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
390                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
392                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
394                 If .Invent.MunicionEqpObjIndex > 0 Then
396                     Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                    End If
        
398                 .Invent.Object(slot).Equipped = 1
400                 .Invent.MunicionEqpObjIndex = .Invent.Object(slot).ObjIndex
402                 .Invent.MunicionEqpSlot = slot

404             Case eOBJType.otArmadura
                
406                 If obj.Ropaje = 0 Then
408                     Call WriteConsoleMsg(UserIndex, "Hay un error con este objeto. Infórmale a un administrador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Nos aseguramos que puede usarla
410                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Or _
                       Not SexoPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Or _
                       Not CheckRazaUsaRopa(UserIndex, .Invent.Object(slot).ObjIndex) Or _
                       Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
                    
412                     Call WriteConsoleMsg(UserIndex, "Tu clase, género, raza o facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
414                 If .Invent.Object(slot).Equipped Then
                    
416                     Call Desequipar(UserIndex, slot)

418                     If .flags.Navegando = 0 Then
                        
420                         If .flags.Montado = 0 Then
422                             Call DarCuerpoDesnudo(UserIndex)
424                             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                            End If

                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
426                 If .Invent.ArmourEqpObjIndex > 0 Then
428                     errordesc = "Armadura 2"
430                     Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
432                     errordesc = "Armadura 3"

                    End If
  
                    'Lo equipa
434                 If Len(obj.CreaGRH) <> 0 Then
436                     .Char.Body_Aura = obj.CreaGRH
438                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))

                    End If
            
440                 .Invent.Object(slot).Equipped = 1
442                 .Invent.ArmourEqpObjIndex = .Invent.Object(slot).ObjIndex
444                 .Invent.ArmourEqpSlot = slot
                            
446                 If .flags.Montado = 0 Then
                
448                     If .flags.Navegando = 0 Then
                        
450                         .Char.Body = obj.Ropaje
                
452                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
454                         .flags.Desnudo = 0
            
                        End If

                    End If
                
456                 If obj.ResistenciaMagica > 0 Then
458                     Call WriteUpdateRM(UserIndex)
                    End If
    
460             Case eOBJType.otCASCO
                
462                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
464                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
466                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
468                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    
                    End If
                
                    'Si esta equipado lo quita
470                 If .Invent.Object(slot).Equipped Then
472                     Call Desequipar(UserIndex, slot)
                
474                     .Char.CascoAnim = NingunCasco
476                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Exit Sub

                    End If
    
                    'Quita el anterior
478                 If .Invent.CascoEqpObjIndex > 0 Then
480                     Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If
            
482                 errordesc = "Casco"

                    'Lo equipa
484                 If Len(obj.CreaGRH) <> 0 Then
486                     .Char.Head_Aura = obj.CreaGRH
488                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
                    End If
            
490                 .Invent.Object(slot).Equipped = 1
492                 .Invent.CascoEqpObjIndex = .Invent.Object(slot).ObjIndex
494                 .Invent.CascoEqpSlot = slot
            
496                 If .flags.Navegando = 0 Then
498                     .Char.CascoAnim = obj.CascoAnim
500                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                
502                 If obj.ResistenciaMagica > 0 Then
504                     Call WriteUpdateRM(UserIndex)
                    End If

506             Case eOBJType.otESCUDO

508                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
510                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
512                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
514                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
516                 If .Invent.Object(slot).Equipped Then
518                     Call Desequipar(UserIndex, slot)
                 
520                     .Char.ShieldAnim = NingunEscudo

522                     If .flags.Montado = 0 Then
524                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
     
                    'Quita el anterior
526                 If .Invent.EscudoEqpObjIndex > 0 Then
528                     Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                    End If
     
                    'Lo equipa
             
530                 If .Invent.Object(slot).ObjIndex = 1700 Or _
                       .Invent.Object(slot).ObjIndex = 1730 Or _
                       .Invent.Object(slot).ObjIndex = 1724 Or _
                       .Invent.Object(slot).ObjIndex = 1717 Or _
                       .Invent.Object(slot).ObjIndex = 1699 Then
             
                    Else

532                     If .Invent.WeaponEqpObjIndex > 0 Then
534                         If ObjData(.Invent.WeaponEqpObjIndex).Proyectil = 1 Then
536                             Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
538                             Call WriteConsoleMsg(UserIndex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                            End If
                        End If

                    End If
            
540                 errordesc = "Escudo"
             
542                 If Len(obj.CreaGRH) <> 0 Then
544                     .Char.Escudo_Aura = obj.CreaGRH
546                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
                    End If

548                 .Invent.Object(slot).Equipped = 1
550                 .Invent.EscudoEqpObjIndex = .Invent.Object(slot).ObjIndex
552                 .Invent.EscudoEqpSlot = slot
                 
554                 If .flags.Navegando = 0 Then
556                     If .flags.Montado = 0 Then
558                         .Char.ShieldAnim = obj.ShieldAnim
560                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                    End If
                
562                 If obj.ResistenciaMagica > 0 Then
564                     Call WriteUpdateRM(UserIndex)
                    End If
                
566             Case eOBJType.otDañoMagico

568                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
570                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

572                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
574                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
576                 If .Invent.Object(slot).Equipped Then
578                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
580                 If .Invent.DañoMagicoEqpSlot > 0 Then
582                     Call Desequipar(UserIndex, .Invent.DañoMagicoEqpSlot)
                    End If
                
584                 .Invent.Object(slot).Equipped = 1
586                 .Invent.DañoMagicoEqpObjIndex = .Invent.Object(slot).ObjIndex
588                 .Invent.DañoMagicoEqpSlot = slot
                
590                 If Len(obj.CreaGRH) <> 0 Then
592                     .Char.DM_Aura = obj.CreaGRH
594                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.DM_Aura, False, 6))
                    End If

596                 Call WriteUpdateDM(UserIndex)
                    
598             Case eOBJType.otResistencia

600                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
602                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

604                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
606                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
608                 If .Invent.Object(slot).Equipped Then
610                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
612                 If .Invent.ResistenciaEqpSlot > 0 Then
614                     Call Desequipar(UserIndex, .Invent.ResistenciaEqpSlot)
                    End If
                
616                 .Invent.Object(slot).Equipped = 1
618                 .Invent.ResistenciaEqpObjIndex = .Invent.Object(slot).ObjIndex
620                 .Invent.ResistenciaEqpSlot = slot
                
622                 If Len(obj.CreaGRH) <> 0 Then
624                     .Char.RM_Aura = obj.CreaGRH
626                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.RM_Aura, False, 7))
                    End If

628                 Call WriteUpdateRM(UserIndex)

            End Select
    
        End With

        'Actualiza
630     Call UpdateUserInv(False, UserIndex, slot)

        Exit Sub
    
ErrHandler:
632     Debug.Print errordesc
634     Call LogError("EquiparInvItem Slot:" & slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description & "- " & errordesc)

End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then
102         CheckRazaUsaRopa = True
            Exit Function

        End If
        
        Dim i As Long
104     For i = 1 To NUMCLASES

106         If ObjData(ItemIndex).RazaProhibida(i) = UserList(UserIndex).raza Then
108             CheckRazaUsaRopa = False
                Exit Function

            End If

110     Next i
        
112     Select Case UserList(UserIndex).raza

            Case eRaza.Humano

114             If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 And ObjData(ItemIndex).RazaDrow = 0 Then
116                 If ObjData(ItemIndex).Ropaje > 0 Then
118                     CheckRazaUsaRopa = True
                        Exit Function

                    End If

                End If

120         Case eRaza.Elfo

122             If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 And ObjData(ItemIndex).RazaDrow = 0 Then
124                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
126         Case eRaza.Orco

128             If ObjData(ItemIndex).RazaEnana = 0 Then
130                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
132         Case eRaza.Drow

134             If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 Then
136                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
138         Case eRaza.Gnomo

140             If ObjData(ItemIndex).RazaEnana > 0 Then
142                 CheckRazaUsaRopa = True
                    Exit Function

                End If
        
144         Case eRaza.Enano

146             If ObjData(ItemIndex).RazaEnana > 0 Then
148                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
        End Select

150     CheckRazaUsaRopa = False

        Exit Function
ErrHandler:
152     Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Public Function CheckRazaTipo(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then

102         CheckRazaTipo = True
            Exit Function

        End If

104     Select Case ObjData(ItemIndex).RazaTipo

            Case 0
106             CheckRazaTipo = True

108         Case 1

110             If UserList(UserIndex).raza = eRaza.Elfo Then
112                 CheckRazaTipo = True
                    Exit Function

                End If
        
114             If UserList(UserIndex).raza = eRaza.Drow Then
116                 CheckRazaTipo = True
                    Exit Function

                End If
        
118             If UserList(UserIndex).raza = eRaza.Humano Then
120                 CheckRazaTipo = True
                    Exit Function

                End If

122         Case 2

124             If UserList(UserIndex).raza = eRaza.Gnomo Then CheckRazaTipo = True
126             If UserList(UserIndex).raza = eRaza.Enano Then CheckRazaTipo = True
                Exit Function

128         Case 3

130             If UserList(UserIndex).raza = eRaza.Orco Then CheckRazaTipo = True
                Exit Function
    
        End Select

        Exit Function
ErrHandler:
132     Call LogError("Error CheckRazaTipo ItemIndex:" & ItemIndex)

End Function

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

Sub UseInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)

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

102         If .Invent.Object(slot).Amount = 0 Then Exit Sub
    
104         obj = ObjData(.Invent.Object(slot).ObjIndex)
    
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
    
134         ObjIndex = .Invent.Object(slot).ObjIndex
136         .flags.TargetObjInvIndex = ObjIndex
138         .flags.TargetObjInvSlot = slot
    
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
            
154                 If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
156                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, .Pos.X, .Pos.Y))
                    Else
158                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, .Pos.X, .Pos.Y))
    
                    End If
            
                    'Quitamos del inv el item
160                 Call QuitarUserInvItem(UserIndex, slot, 1)
            
162                 Call UpdateUserInv(False, UserIndex, slot)
    
164             Case eOBJType.otGuita
    
166                 If .flags.Muerto = 1 Then
168                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
170                 .Stats.GLD = .Stats.GLD + .Invent.Object(slot).Amount
172                 .Invent.Object(slot).Amount = 0
174                 .Invent.Object(slot).ObjIndex = 0
176                 .Invent.NroItems = .Invent.NroItems - 1
            
178                 Call UpdateUserInv(False, UserIndex, slot)
180                 Call WriteUpdateGold(UserIndex)
            
182             Case eOBJType.otWeapon
    
184                 If .flags.Muerto = 1 Then
186                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
188                 If Not .Stats.MinSta > 0 Then
190                     Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
192                 If ObjData(ObjIndex).Proyectil = 1 Then
                        'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
194                     Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                    Else
    
196                     If .flags.TargetObj = Leña Then
198                         If .Invent.Object(slot).ObjIndex = DAGA Then
200                             Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
    
                            End If
    
                        End If
    
                    End If
            
                    'REVISAR LADDER
                    'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
202                 If .Invent.Object(slot).Equipped = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
204             Case eOBJType.otHerramientas
    
206                 If .flags.Muerto = 1 Then
208                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
210                 If Not .Stats.MinSta > 0 Then
212                     Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
                    'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
214                 If .Invent.Object(slot).Equipped = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
216                     Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
218                 Select Case obj.Subtipo
                    
                        Case 1, 2  ' Herramientas del Pescador - Caña y Red
220                         Call WriteWorkRequestTarget(UserIndex, eSkill.Pescar)
                    
222                     Case 3     ' Herramientas de Alquimia - Tijeras
224                         Call WriteWorkRequestTarget(UserIndex, eSkill.Alquimia)
                    
226                     Case 4     ' Herramientas de Alquimia - Olla
228                         Call EnivarObjConstruiblesAlquimia(UserIndex)
230                         Call WriteShowAlquimiaForm(UserIndex)
                    
232                     Case 5     ' Herramientas de Carpinteria - Serrucho
234                         Call EnivarObjConstruibles(UserIndex)
236                         Call WriteShowCarpenterForm(UserIndex)
                    
238                     Case 6     ' Herramientas de Tala - Hacha
240                         Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)
    
242                     Case 7     ' Herramientas de Herrero - Martillo
244                         Call WriteConsoleMsg(UserIndex, "Debes hacer click derecho sobre el yunque.", FontTypeNames.FONTTYPE_INFOIAO)
    
246                     Case 8     ' Herramientas de Mineria - Piquete
248                         Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
                    
250                     Case 9     ' Herramientas de Sastreria - Costurero
252                         Call EnivarObjConstruiblesSastre(UserIndex)
254                         Call WriteShowSastreForm(UserIndex)
    
                    End Select
        
256             Case eOBJType.otPociones
    
258                 If .flags.Muerto = 1 Then
260                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
262                 .flags.TomoPocion = True
264                 .flags.TipoPocion = obj.TipoPocion
                    
                    Dim CabezaFinal  As Integer
    
                    Dim CabezaActual As Integer
    
266                 Select Case .flags.TipoPocion
            
                        Case 1 'Modif la agilidad
268                         .flags.DuracionEfecto = obj.DuracionEfecto
            
                            'Usa el item
270                         .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(.Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(eAtributos.Agilidad) * 2)
                    
272                         Call WriteFYA(UserIndex)
                    
                            'Quitamos del inv el item
274                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
276                         If obj.Snd1 <> 0 Then
278                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
280                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
            
282                     Case 2 'Modif la fuerza
284                         .flags.DuracionEfecto = obj.DuracionEfecto
            
                            'Usa el item
286                         .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(.Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(eAtributos.Fuerza) * 2)
                    
                            'Quitamos del inv el item
288                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
290                         If obj.Snd1 <> 0 Then
292                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
294                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
296                         Call WriteFYA(UserIndex)
    
298                     Case 3 'Pocion roja, restaura HP
                    
                            'Usa el item
300                         .Stats.MinHp = .Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)
    
302                         If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                    
                            'Quitamos del inv el item
304                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
306                         If obj.Snd1 <> 0 Then
308                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                            Else
310                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
                
312                     Case 4 'Pocion azul, restaura MANA
                
                            Dim porcentajeRec As Byte
314                         porcentajeRec = obj.Porcentaje
                    
                            'Usa el item
316                         .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, porcentajeRec)
    
318                         If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                    
                            'Quitamos del inv el item
320                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
322                         If obj.Snd1 <> 0 Then
324                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                            Else
326                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
                    
328                     Case 5 ' Pocion violeta
    
330                         If .flags.Envenenado > 0 Then
332                             .flags.Envenenado = 0
334                             Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                                'Quitamos del inv el item
336                             Call QuitarUserInvItem(UserIndex, slot, 1)
    
338                             If obj.Snd1 <> 0 Then
340                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                                Else
342                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                                End If
    
                            Else
344                             Call WriteConsoleMsg(UserIndex, "¡No te encuentras envenenado!", FontTypeNames.FONTTYPE_INFO)
    
                            End If
                    
346                     Case 6  ' Remueve Parálisis
    
348                         If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
350                             If .flags.Paralizado = 1 Then
352                                 .flags.Paralizado = 0
354                                 Call WriteParalizeOK(UserIndex)
    
                                End If
                            
356                             If .flags.Inmovilizado = 1 Then
358                                 .Counters.Inmovilizado = 0
360                                 .flags.Inmovilizado = 0
362                                 Call WriteInmovilizaOK(UserIndex)
    
                                End If
                            
                            
                            
364                             Call QuitarUserInvItem(UserIndex, slot, 1)
    
366                             If obj.Snd1 <> 0 Then
368                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
                                Else
370                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(255, .Pos.X, .Pos.Y))
    
                                End If
    
372                             Call WriteConsoleMsg(UserIndex, "Te has removido la paralizis.", FontTypeNames.FONTTYPE_INFOIAO)
                            Else
374                             Call WriteConsoleMsg(UserIndex, "No estas paralizado.", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
                    
376                     Case 7  ' Pocion Naranja
378                         .Stats.MinSta = .Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)
    
380                         If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                        
                            'Quitamos del inv el item
382                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
384                         If obj.Snd1 <> 0 Then
386                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                            Else
388                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
390                     Case 8  ' Pocion cambio cara
    
392                         Select Case .genero
    
                                Case eGenero.Hombre
    
394                                 Select Case .raza
    
                                        Case eRaza.Humano
396                                         CabezaFinal = RandomNumber(1, 40)
    
398                                     Case eRaza.Elfo
400                                         CabezaFinal = RandomNumber(101, 132)
    
402                                     Case eRaza.Drow
404                                         CabezaFinal = RandomNumber(201, 229)
    
406                                     Case eRaza.Enano
408                                         CabezaFinal = RandomNumber(301, 329)
    
410                                     Case eRaza.Gnomo
412                                         CabezaFinal = RandomNumber(401, 429)
    
414                                     Case eRaza.Orco
416                                         CabezaFinal = RandomNumber(501, 529)
    
                                    End Select
    
418                             Case eGenero.Mujer
    
420                                 Select Case .raza
    
                                        Case eRaza.Humano
422                                         CabezaFinal = RandomNumber(50, 80)
    
424                                     Case eRaza.Elfo
426                                         CabezaFinal = RandomNumber(150, 179)
    
428                                     Case eRaza.Drow
430                                         CabezaFinal = RandomNumber(250, 279)
    
432                                     Case eRaza.Gnomo
434                                         CabezaFinal = RandomNumber(350, 379)
    
436                                     Case eRaza.Enano
438                                         CabezaFinal = RandomNumber(450, 479)
    
440                                     Case eRaza.Orco
442                                         CabezaFinal = RandomNumber(550, 579)
    
                                    End Select
    
                            End Select
                
444                         .Char.Head = CabezaFinal
446                         .OrigChar.Head = CabezaFinal
448                         Call ChangeUserChar(UserIndex, .Char.Body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                            'Quitamos del inv el item
450                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 102, 0))
    
452                         If CabezaActual <> CabezaFinal Then
454                             Call QuitarUserInvItem(UserIndex, slot, 1)
                            Else
456                             Call WriteConsoleMsg(UserIndex, "¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
458                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        
460                     Case 9  ' Pocion sexo
        
462                         Select Case .genero
    
                                Case eGenero.Hombre
464                                 .genero = eGenero.Mujer
                        
466                             Case eGenero.Mujer
468                                 .genero = eGenero.Hombre
                        
                            End Select
                
470                         Select Case .genero
    
                                Case eGenero.Hombre
    
472                                 Select Case .raza
    
                                        Case eRaza.Humano
474                                         CabezaFinal = RandomNumber(1, 40)
    
476                                     Case eRaza.Elfo
478                                         CabezaFinal = RandomNumber(101, 132)
    
480                                     Case eRaza.Drow
482                                         CabezaFinal = RandomNumber(201, 229)
    
484                                     Case eRaza.Enano
486                                         CabezaFinal = RandomNumber(301, 329)
    
488                                     Case eRaza.Gnomo
490                                         CabezaFinal = RandomNumber(401, 429)
    
492                                     Case eRaza.Orco
494                                         CabezaFinal = RandomNumber(501, 529)
    
                                    End Select
    
496                             Case eGenero.Mujer
    
498                                 Select Case .raza
    
                                        Case eRaza.Humano
500                                         CabezaFinal = RandomNumber(50, 80)
    
502                                     Case eRaza.Elfo
504                                         CabezaFinal = RandomNumber(150, 179)
    
506                                     Case eRaza.Drow
508                                         CabezaFinal = RandomNumber(250, 279)
    
510                                     Case eRaza.Gnomo
512                                         CabezaFinal = RandomNumber(350, 379)
    
514                                     Case eRaza.Enano
516                                         CabezaFinal = RandomNumber(450, 479)
    
518                                     Case eRaza.Orco
520                                         CabezaFinal = RandomNumber(550, 579)
    
                                    End Select
    
                            End Select
                
522                         .Char.Head = CabezaFinal
524                         .OrigChar.Head = CabezaFinal
526                         Call ChangeUserChar(UserIndex, .Char.Body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                            'Quitamos del inv el item
528                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 102, 0))
530                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
532                         If obj.Snd1 <> 0 Then
534                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            Else
536                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
                    
538                     Case 10  ' Invisibilidad
                
540                         If .flags.invisible = 0 Then
542                             .flags.invisible = 1
544                             .Counters.Invisibilidad = obj.DuracionEfecto
546                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
548                             Call WriteContadores(UserIndex)
550                             Call QuitarUserInvItem(UserIndex, slot, 1)
    
552                             If obj.Snd1 <> 0 Then
554                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                                Else
556                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("123", .Pos.X, .Pos.Y))
    
                                End If
    
558                             Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras...", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            
                            Else
560                             Call WriteConsoleMsg(UserIndex, "Ya estas invisible.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                Exit Sub
    
                            End If
                        
562                     Case 11  ' Experiencia
    
                            Dim HR   As Integer
    
                            Dim MS   As Integer
    
                            Dim SS   As Integer
    
                            Dim secs As Integer
    
564                         If .flags.ScrollExp = 1 Then
566                             .flags.ScrollExp = obj.CuantoAumento
568                             .Counters.ScrollExperiencia = obj.DuracionEfecto
570                             Call QuitarUserInvItem(UserIndex, slot, 1)
                            
572                             secs = obj.DuracionEfecto
574                             HR = secs \ 3600
576                             MS = (secs Mod 3600) \ 60
578                             SS = (secs Mod 3600) Mod 60
    
580                             If SS > 9 Then
582                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                                Else
584                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
    
                                End If
    
                            Else
586                             Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                                Exit Sub
    
                            End If
    
588                         Call WriteContadores(UserIndex)
    
590                         If obj.Snd1 <> 0 Then
592                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
594                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
596                     Case 12  ' Oro
                
598                         If .flags.ScrollOro = 1 Then
600                             .flags.ScrollOro = obj.CuantoAumento
602                             .Counters.ScrollOro = obj.DuracionEfecto
604                             Call QuitarUserInvItem(UserIndex, slot, 1)
606                             secs = obj.DuracionEfecto
608                             HR = secs \ 3600
610                             MS = (secs Mod 3600) \ 60
612                             SS = (secs Mod 3600) Mod 60
    
614                             If SS > 9 Then
616                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                                Else
618                                 Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
    
                                End If
                            
                            Else
620                             Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                                Exit Sub
    
                            End If
    
622                         Call WriteContadores(UserIndex)
    
624                         If obj.Snd1 <> 0 Then
626                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
628                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
630                     Case 13
                    
632                         Call QuitarUserInvItem(UserIndex, slot, 1)
634                         .flags.Envenenado = 0
636                         .flags.Incinerado = 0
                        
638                         If .flags.Inmovilizado = 1 Then
640                             .Counters.Inmovilizado = 0
642                             .flags.Inmovilizado = 0
644                             Call WriteInmovilizaOK(UserIndex)
                            
    
                            End If
                        
646                         If .flags.Paralizado = 1 Then
648                             .flags.Paralizado = 0
650                             Call WriteParalizeOK(UserIndex)
                            
    
                            End If
                        
652                         If .flags.Ceguera = 1 Then
654                             .flags.Ceguera = 0
656                             Call WriteBlindNoMore(UserIndex)
                            
    
                            End If
                        
658                         If .flags.Maldicion = 1 Then
660                             .flags.Maldicion = 0
662                             .Counters.Maldicion = 0
    
                            End If
                        
664                         .Stats.MinSta = .Stats.MaxSta
666                         .Stats.MinAGU = .Stats.MaxAGU
668                         .Stats.MinMAN = .Stats.MaxMAN
670                         .Stats.MinHp = .Stats.MaxHp
672                         .Stats.MinHam = .Stats.MaxHam
                        
674                         .flags.Hambre = 0
676                         .flags.Sed = 0
                        
678                         Call WriteUpdateHungerAndThirst(UserIndex)
680                         Call WriteConsoleMsg(UserIndex, "Donador> Te sentis sano y lleno.", FontTypeNames.FONTTYPE_WARNING)
    
682                         If obj.Snd1 <> 0 Then
684                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
686                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
688                     Case 14
                    
690                         If .flags.BattleModo = 1 Then
692                             Call WriteConsoleMsg(UserIndex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                                Exit Sub
    
                            End If
                        
694                         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
696                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                        
                            Dim Map     As Integer
    
                            Dim X       As Byte
    
                            Dim Y       As Byte
    
                            Dim DeDonde As WorldPos
    
698                         Call QuitarUserInvItem(UserIndex, slot, 1)
                
700                         Select Case .Hogar
    
                                Case eCiudad.cUllathorpe
702                                 DeDonde = Ullathorpe
                                
704                             Case eCiudad.cNix
706                                 DeDonde = Nix
                    
708                             Case eCiudad.cBanderbill
710                                 DeDonde = Banderbill
                            
712                             Case eCiudad.cLindos
714                                 DeDonde = Lindos
                                
716                             Case eCiudad.cArghal
718                                 DeDonde = Arghal
                                
720                             Case eCiudad.CHillidan
722                                 DeDonde = Hillidan
                                
724                             Case Else
726                                 DeDonde = Ullathorpe
    
                            End Select
                        
728                         Map = DeDonde.Map
730                         X = DeDonde.X
732                         Y = DeDonde.Y
                        
734                         Call FindLegalPos(UserIndex, Map, X, Y)
736                         Call WarpUserChar(UserIndex, Map, X, Y, True)
738                         Call WriteConsoleMsg(UserIndex, "Ya estas a salvo...", FontTypeNames.FONTTYPE_WARNING)
    
740                         If obj.Snd1 <> 0 Then
742                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                            
                            Else
744                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                            End If
    
746                     Case 15  ' Aliento de sirena
                            
748                         If .Counters.Oxigeno >= 3540 Then
                            
750                             Call WriteConsoleMsg(UserIndex, "No podes acumular más de 59 minutos de oxigeno.", FontTypeNames.FONTTYPE_INFOIAO)
752                             secs = .Counters.Oxigeno
754                             HR = secs \ 3600
756                             MS = (secs Mod 3600) \ 60
758                             SS = (secs Mod 3600) Mod 60
    
760                             If SS > 9 Then
762                                 Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                                Else
764                                 Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)
    
                                End If
    
                            Else
                                
766                             .Counters.Oxigeno = .Counters.Oxigeno + obj.DuracionEfecto
768                             Call QuitarUserInvItem(UserIndex, slot, 1)
                                
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
                                
770                             .flags.Ahogandose = 0
772                             Call WriteOxigeno(UserIndex)
                                
774                             Call WriteContadores(UserIndex)
    
776                             If obj.Snd1 <> 0 Then
778                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                                Else
780                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                                End If
    
                            End If
    
782                     Case 16 ' Divorcio
    
784                         If .flags.Casado = 1 Then
    
                                Dim tUser As Integer
    
                                '.flags.Pareja
786                             tUser = NameIndex(.flags.Pareja)
788                             Call QuitarUserInvItem(UserIndex, slot, 1)
                            
790                             If tUser <= 0 Then
    
                                    Dim FileUser As String
    
792                                 FileUser = CharPath & UCase$(.flags.Pareja) & ".chr"
                                    'Call WriteVar(FileUser, "FLAGS", "CASADO", 0)
                                    'Call WriteVar(FileUser, "FLAGS", "PAREJA", "")
794                                 .flags.Casado = 0
796                                 .flags.Pareja = ""
798                                 Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
800                                 .MENSAJEINFORMACION = .name & " se ha divorciado de ti."
    
                                Else
802                                 UserList(tUser).flags.Casado = 0
804                                 UserList(tUser).flags.Pareja = ""
806                                 .flags.Casado = 0
808                                 .flags.Pareja = ""
810                                 Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
812                                 Call WriteConsoleMsg(tUser, .name & " se ha divorciado de ti.", FontTypeNames.FONTTYPE_INFOIAO)
                                
                                End If
    
814                             If obj.Snd1 <> 0 Then
816                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                
                                Else
818                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                                End If
                        
                            Else
820                             Call WriteConsoleMsg(UserIndex, "No estas casado.", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
822                     Case 17 'Cara legendaria
    
824                         Select Case .genero
    
                                Case eGenero.Hombre
    
826                                 Select Case .raza
    
                                        Case eRaza.Humano
828                                         CabezaFinal = RandomNumber(684, 686)
    
830                                     Case eRaza.Elfo
832                                         CabezaFinal = RandomNumber(690, 692)
    
834                                     Case eRaza.Drow
836                                         CabezaFinal = RandomNumber(696, 698)
    
838                                     Case eRaza.Enano
840                                         CabezaFinal = RandomNumber(702, 704)
    
842                                     Case eRaza.Gnomo
844                                         CabezaFinal = RandomNumber(708, 710)
    
846                                     Case eRaza.Orco
848                                         CabezaFinal = RandomNumber(714, 716)
    
                                    End Select
    
850                             Case eGenero.Mujer
    
852                                 Select Case .raza
    
                                        Case eRaza.Humano
854                                         CabezaFinal = RandomNumber(687, 689)
    
856                                     Case eRaza.Elfo
858                                         CabezaFinal = RandomNumber(693, 695)
    
860                                     Case eRaza.Drow
862                                         CabezaFinal = RandomNumber(699, 701)
    
864                                     Case eRaza.Gnomo
866                                         CabezaFinal = RandomNumber(705, 707)
    
868                                     Case eRaza.Enano
870                                         CabezaFinal = RandomNumber(711, 713)
    
872                                     Case eRaza.Orco
874                                         CabezaFinal = RandomNumber(717, 719)
    
                                    End Select
    
                            End Select
    
876                         CabezaActual = .OrigChar.Head
                            
878                         .Char.Head = CabezaFinal
880                         .OrigChar.Head = CabezaFinal
882                         Call ChangeUserChar(UserIndex, .Char.Body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    
                            'Quitamos del inv el item
884                         If CabezaActual <> CabezaFinal Then
886                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 102, 0))
888                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
890                             Call QuitarUserInvItem(UserIndex, slot, 1)
                            Else
892                             Call WriteConsoleMsg(UserIndex, "¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!", FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
894                     Case 18  ' tan solo crea una particula por determinado tiempo
    
                            Dim Particula           As Integer
    
                            Dim Tiempo              As Long
    
                            Dim ParticulaPermanente As Byte
    
                            Dim sobrechar           As Byte
    
896                         If obj.CreaParticula <> "" Then
898                             Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
900                             Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
902                             ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
904                             sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                                
906                             If ParticulaPermanente = 1 Then
908                                 .Char.ParticulaFx = Particula
910                                 .Char.loops = Tiempo
    
                                End If
                                
912                             If sobrechar = 1 Then
914                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(.Pos.X, .Pos.Y, Particula, Tiempo))
                                Else
                                
916                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, Particula, Tiempo, False))
    
                                    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(.Pos.x, .Pos.Y, Particula, Tiempo))
                                End If
    
                            End If
                            
918                         If obj.CreaFX <> 0 Then
920                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, .Pos.X, .Pos.Y))
                                
                                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, obj.CreaFX, 0))
                                ' PrepareMessageCreateFX
                            End If
                            
922                         If obj.Snd1 <> 0 Then
924                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
    
                            End If
                            
926                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
928                     Case 19 ' Reseteo de skill
    
                            Dim s As Byte
                    
930                         If .Stats.UserSkills(eSkill.Liderazgo) >= 80 Then
932                             Call WriteConsoleMsg(UserIndex, "Has fundado un clan, no podes resetar tus skills. ", FontTypeNames.FONTTYPE_INFOIAO)
                                Exit Sub
    
                            End If
                        
934                         For s = 1 To NUMSKILLS
936                             .Stats.UserSkills(s) = 0
938                         Next s
                        
                            Dim SkillLibres As Integer
                        
940                         SkillLibres = 5
942                         SkillLibres = SkillLibres + (5 * .Stats.ELV)
                         
944                         .Stats.SkillPts = SkillLibres
946                         Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                        
948                         Call WriteConsoleMsg(UserIndex, "Tus skills han sido reseteados.", FontTypeNames.FONTTYPE_INFOIAO)
950                         Call QuitarUserInvItem(UserIndex, slot, 1)
    
952                     Case 20
                    
954                         If .Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
956                             .Stats.InventLevel = .Stats.InventLevel + 1
958                             .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
960                             Call WriteInventoryUnlockSlots(UserIndex)
962                             Call WriteConsoleMsg(UserIndex, "Has aumentado el espacio de tu inventario!", FontTypeNames.FONTTYPE_INFO)
964                             Call QuitarUserInvItem(UserIndex, slot, 1)
                            Else
966                             Call WriteConsoleMsg(UserIndex, "Ya has desbloqueado todos los casilleros disponibles.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                    
                    End Select
    
968                 Call WriteUpdateUserStats(UserIndex)
970                 Call UpdateUserInv(False, UserIndex, slot)
    
972             Case eOBJType.otBebidas
    
974                 If .flags.Muerto = 1 Then
976                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
978                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
    
980                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
982                 .flags.Sed = 0
984                 Call WriteUpdateHungerAndThirst(UserIndex)
            
                    'Quitamos del inv el item
986                 Call QuitarUserInvItem(UserIndex, slot, 1)
            
988                 If obj.Snd1 <> 0 Then
990                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                
                    Else
992                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
    
                     End If
            
994                 Call UpdateUserInv(False, UserIndex, slot)
            
996             Case eOBJType.OtCofre
    
998                 If .flags.Muerto = 1 Then
1000                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
                       'Quitamos del inv el item
1002                 Call QuitarUserInvItem(UserIndex, slot, 1)
1004                 Call UpdateUserInv(False, UserIndex, slot)
            
1006                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(.name & " ha abierto un " & obj.name & " y obtuvo...", FontTypeNames.FONTTYPE_New_DONADOR))
            
1008                 If obj.Snd1 <> 0 Then
1010                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                       End If
            
1012                 If obj.CreaFX <> 0 Then
1014                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, obj.CreaFX, 0))
                       End If
            
                       Dim i As Byte
    
1016                 If obj.Subtipo = 1 Then
    
1018                     For i = 1 To obj.CantItem

1020                        If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                
1022                             If (.flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) Then
1024                                 Call TirarItemAlPiso(.Pos, obj.Item(i))
                                  End If
                                
                              End If
                            
1026                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))

1028                     Next i
            
                       Else
            
1030                     For i = 1 To obj.CantEntrega
    
                               Dim indexobj As Byte
1032                            indexobj = RandomNumber(1, obj.CantItem)
                
                               Dim index As obj
    
1034                         index.ObjIndex = obj.Item(indexobj).ObjIndex
1036                         index.Amount = obj.Item(indexobj).Amount
    
1038                         If Not MeterItemEnInventario(UserIndex, index) Then

1040                            If (.flags.Privilegios And (PlayerType.user Or PlayerType.Dios Or PlayerType.Admin)) Then
1042                                 Call TirarItemAlPiso(.Pos, index)
                                  End If
                                
                               End If

1044                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(index.ObjIndex).name & " (" & index.Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
1046                     Next i
    
                       End If
        
1048             Case eOBJType.otLlaves
1050                 Call WriteConsoleMsg(UserIndex, "Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.", FontTypeNames.FONTTYPE_INFO)
        
1052             Case eOBJType.otBotellaVacia
    
1054                 If .flags.Muerto = 1 Then
1056                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1058                 If (MapData(.Pos.Map, .flags.TargetX, .flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
1060                     Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1062                 MiObj.Amount = 1
1064                 MiObj.ObjIndex = ObjData(.Invent.Object(slot).ObjIndex).IndexAbierta

1066                 Call QuitarUserInvItem(UserIndex, slot, 1)
    
1068                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
1070                     Call TirarItemAlPiso(.Pos, MiObj)
                       End If
            
1072                 Call UpdateUserInv(False, UserIndex, slot)
        
1074             Case eOBJType.otBotellaLlena
    
1076                 If .flags.Muerto = 1 Then
1078                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1080                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
    
1082                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
1084                 .flags.Sed = 0
1086                 Call WriteUpdateHungerAndThirst(UserIndex)
1088                 MiObj.Amount = 1
1090                 MiObj.ObjIndex = ObjData(.Invent.Object(slot).ObjIndex).IndexCerrada
1092                 Call QuitarUserInvItem(UserIndex, slot, 1)
    
1094                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
1096                     Call TirarItemAlPiso(.Pos, MiObj)
    
                       End If
            
1098                 Call UpdateUserInv(False, UserIndex, slot)
        
1100             Case eOBJType.otPergaminos
    
1102                 If .flags.Muerto = 1 Then
1104                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
                       'Call LogError(.Name & " intento aprender el hechizo " & ObjData(.Invent.Object(slot).ObjIndex).HechizoIndex)
            
1106                 If ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
    
                           'If .Stats.MaxMAN > 0 Then
1108                     If .flags.Hambre = 0 And .flags.Sed = 0 Then
1110                         Call AgregarHechizo(UserIndex, slot)
1112                         Call UpdateUserInv(False, UserIndex, slot)
                               ' Call LogError(.Name & " lo aprendio.")
                           Else
1114                         Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
    
                           End If
    
                           ' Else
                           '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_WARNING)
                           'End If
                       Else
                 
1116                     Call WriteConsoleMsg(UserIndex, "Por mas que lo intentas, no podés comprender el manuescrito.", FontTypeNames.FONTTYPE_INFO)
       
                       End If
            
1118             Case eOBJType.otMinerales
    
1120                 If .flags.Muerto = 1 Then
1122                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1124                 Call WriteWorkRequestTarget(UserIndex, FundirMetal)
           
1126             Case eOBJType.otInstrumentos
    
1128                 If .flags.Muerto = 1 Then
1130                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1132                 If obj.Real Then '¿Es el Cuerno Real?
1134                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1136                         If MapInfo(.Pos.Map).Seguro = 1 Then
1138                             Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                                   Exit Sub
    
                               End If
    
1140                         Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                               Exit Sub
                           Else
1142                         Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                               Exit Sub
    
                           End If
    
1144                 ElseIf obj.Caos Then '¿Es el Cuerno Legión?
    
1146                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1148                         If MapInfo(.Pos.Map).Seguro = 1 Then
1150                             Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                                   Exit Sub
    
                               End If
    
1152                         Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                               Exit Sub
                           Else
1154                         Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                               Exit Sub
    
                           End If
    
                       End If
    
                       'Si llega aca es porque es o Laud o Tambor o Flauta
1156                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
           
1158             Case eOBJType.otBarcos
                
                     ' Piratas y trabajadores navegan al nivel 20
1160                 If .clase = eClass.Trabajador Or .clase = eClass.Pirat Then
1162                     If .Stats.ELV < 20 Then
1164                         Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub
                         End If
                    
                     ' Nivel mínimo 25 para navegar, si no sos pirata ni trabajador
1166                ElseIf .Stats.ELV < 25 Then
1168                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                         Exit Sub
                     End If

1170                If .flags.Navegando = 0 Then
1172                    If LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False) Then
1174                        Call DoNavega(UserIndex, obj, slot)
                          Else
1176                        Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
                          End If
                    
                      Else
1178                     If .Invent.BarcoObjIndex <> .Invent.Object(slot).ObjIndex Then
1180                        Call DoNavega(UserIndex, obj, slot)
                          Else
1182                        If LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, False, True) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, False, True) Then
1184                            Call DoNavega(UserIndex, obj, slot)
                              Else
1186                            Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte a la costa para dejar la barca!", FontTypeNames.FONTTYPE_INFO)
                              End If
                          End If
                      End If
            
1188             Case eOBJType.otMonturas
                       'Verifica todo lo que requiere la montura
        
1190                 If .flags.Muerto = 1 Then
1192                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡Estas muerto! Los fantasmas no pueden montar.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
                
1194                 If .flags.Navegando = 1 Then
1196                     Call WriteConsoleMsg(UserIndex, "Debes dejar de navegar para poder montarté.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1198                 If MapInfo(.Pos.Map).zone = "DUNGEON" Then
1200                     Call WriteConsoleMsg(UserIndex, "No podes cabalgar dentro de un dungeon.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1202                 Call DoMontar(UserIndex, obj, slot)
    
1204             Case eOBJType.OtDonador
    
1206                 Select Case obj.Subtipo
    
                           Case 1
                
1208                         If .Counters.Pena <> 0 Then
1210                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                                   Exit Sub
    
                               End If
                    
1212                         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
1214                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                                   Exit Sub
    
                               End If
                
1216                         Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1218                         Call WriteConsoleMsg(UserIndex, "Has viajado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
1220                         Call QuitarUserInvItem(UserIndex, slot, 1)
1222                         Call UpdateUserInv(False, UserIndex, slot)
                    
1224                     Case 2
    
1226                         If DonadorCheck(.Cuenta) = 0 Then
1228                             Call DonadorTiempo(.Cuenta, CLng(obj.CuantoAumento))
1230                             Call WriteConsoleMsg(UserIndex, "Donación> Se han agregado " & obj.CuantoAumento & " dias de donador a tu cuenta. Relogea tu personaje para empezar a disfrutar la experiencia.", FontTypeNames.FONTTYPE_WARNING)
1232                             Call QuitarUserInvItem(UserIndex, slot, 1)
1234                             Call UpdateUserInv(False, UserIndex, slot)
                               Else
1236                             Call DonadorTiempo(.Cuenta, CLng(obj.CuantoAumento))
1238                             Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & CLng(obj.CuantoAumento) & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
1240                             .donador.activo = 1
1242                             Call QuitarUserInvItem(UserIndex, slot, 1)
1244                             Call UpdateUserInv(False, UserIndex, slot)
    
                                   'Call WriteConsoleMsg(UserIndex, "Donación> Debes esperar a que finalice el periodo existente para renovar tu suscripción.", FontTypeNames.FONTTYPE_INFOIAO)
                               End If
    
1246                     Case 3
1248                         Call AgregarCreditosDonador(.Cuenta, CLng(obj.CuantoAumento))
1250                         Call WriteConsoleMsg(UserIndex, "Donación> Tu credito ahora es de " & CreditosDonadorCheck(.Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
1252                         Call QuitarUserInvItem(UserIndex, slot, 1)
1254                         Call UpdateUserInv(False, UserIndex, slot)
    
                       End Select
         
1256             Case eOBJType.otpasajes
    
1258                 If .flags.Muerto = 1 Then
1260                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1262                 If .flags.TargetNpcTipo <> Pirata Then
1264                     Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1266                 If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
1268                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1270                 If .Pos.Map <> obj.DesdeMap Then
                           Rem  Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aquí! Largate!", FontTypeNames.FONTTYPE_INFO)
1272                     Call WriteChatOverHead(UserIndex, "El pasaje no lo compraste aquí! Largate!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                           Exit Sub
    
                       End If
            
1274                 If Not MapaValido(obj.HastaMap) Then
                           Rem Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
1276                     Call WriteChatOverHead(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                           Exit Sub
    
                       End If
    
1278                 If obj.NecesitaNave > 0 Then
1280                     If .Stats.UserSkills(eSkill.Navegacion) < 80 Then
                               Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
1282                         Call WriteChatOverHead(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                               Exit Sub
    
                           End If
    
                       End If
                
1284                 Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1286                 Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
1288                 .Stats.MinAGU = 0
1290                 .Stats.MinHam = 0
1292                 .flags.Sed = 1
1294                 .flags.Hambre = 1
1296                 Call WriteUpdateHungerAndThirst(UserIndex)
1298                 Call QuitarUserInvItem(UserIndex, slot, 1)
1300                 Call UpdateUserInv(False, UserIndex, slot)
            
1302             Case eOBJType.otRunas
        
1304                 If .Counters.Pena <> 0 Then
1306                     Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1308                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
1310                     Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1312                 If .flags.BattleModo = 1 Then
1314                     Call WriteConsoleMsg(UserIndex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                           Exit Sub
    
                       End If
            
1316                 If MapInfo(.Pos.Map).Seguro = 0 And .flags.Muerto = 0 Then
1318                     Call WriteConsoleMsg(UserIndex, "Solo podes usar tu runa en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1320                 If .Accion.AccionPendiente Then
                           Exit Sub
    
                       End If
            
1322                 Select Case ObjData(ObjIndex).TipoRuna
            
                           Case 1, 2
    
1324                         If .donador.activo = 0 And Not EsGM(UserIndex) Then ' Donador no espera tiempo
1326                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Runa, 400, False))
1328                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 350, Accion_Barra.Runa))
                               Else
1330                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Runa, 50, False))
1332                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 100, Accion_Barra.Runa))
    
                               End If
    
1334                         .Accion.Particula = ParticulasIndex.Runa
1336                         .Accion.AccionPendiente = True
1338                         .Accion.TipoAccion = Accion_Barra.Runa
1340                         .Accion.RunaObj = ObjIndex
1342                         .Accion.ObjSlot = slot
                
1344                     Case 3
            
                               Dim parejaindex As Integer
    
1346                         If Not .flags.BattleModo Then
                    
                                   'If .donador.activo = 1 Then
1348                             If MapInfo(.Pos.Map).Seguro = 1 Then
1350                                 If .flags.Casado = 1 Then
1352                                     parejaindex = NameIndex(.flags.Pareja)
                            
1354                                     If parejaindex > 0 Then
1356                                         If UserList(parejaindex).flags.BattleModo = 0 Then
1358                                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, ParticulasIndex.Runa, 600, False))
1360                                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, 600, Accion_Barra.GoToPareja))
1362                                             .Accion.AccionPendiente = True
1364                                             .Accion.Particula = ParticulasIndex.Runa
1366                                             .Accion.TipoAccion = Accion_Barra.GoToPareja
                                               Else
1368                                             Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)
    
                                               End If
                                    
                                           Else
1370                                         Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)
    
                                           End If
    
                                       Else
1372                                     Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)
    
                                       End If
    
                                   Else
1374                                 Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)
    
                                   End If
                    
                                   ' Else
                                   '  Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                                   ' End If
                               Else
1376                             Call WriteConsoleMsg(UserIndex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
            
                               End If
        
                       End Select
            
1378             Case eOBJType.otmapa
1380                 Call WriteShowFrmMapa(UserIndex)
            
               End Select
             
          End With

          Exit Sub

hErr:
1382    LogError "Error en useinvitem Usuario: " & UserList(UserIndex).name & " item:" & obj.name & " index: " & UserList(UserIndex).Invent.Object(slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmasConstruibles_Err
        

100     Call WriteBlacksmithWeapons(UserIndex)

        
        Exit Sub

EnivarArmasConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.EnivarArmasConstruibles", Erl)
104     Resume Next
        
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruibles_Err
        

100     Call WriteCarpenterObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruibles", Erl)
104     Resume Next
        
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesAlquimia_Err
        

100     Call WriteAlquimistaObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesAlquimia_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)
104     Resume Next
        
End Sub

Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesSastre_Err
        

100     Call WriteSastreObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesSastre_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)
104     Resume Next
        
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmadurasConstruibles_Err
        

100     Call WriteBlacksmithArmors(UserIndex)

        
        Exit Sub

EnivarArmadurasConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.EnivarArmadurasConstruibles", Erl)
104     Resume Next
        
End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
        
        On Error GoTo ItemSeCae_Err
        

100     ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And ObjData(index).OBJType <> eOBJType.otLlaves And ObjData(index).OBJType <> eOBJType.otBarcos And ObjData(index).OBJType <> eOBJType.otMonturas And ObjData(index).NoSeCae = 0 And Not ObjData(index).Intirable = 1 And Not ObjData(index).Destruye = 1 And ObjData(index).donador = 0 And Not ObjData(index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.ItemSeCae", Erl)
104     Resume Next
        
End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo PirataCaeItem_Err

100     With UserList(UserIndex)
    
102         If .clase = eClass.Pirat Then

                ' Si no está navegando, se caen los items
104             If .Invent.BarcoObjIndex > 0 Then
            
                    ' El pirata con galera no pierde los últimos 6 * (cada 10 niveles; max 1) slots
106                 If ObjData(.Invent.BarcoObjIndex).Ropaje = iGalera Then
                
108                     If slot > .CurrentInventorySlots - 6 * min(.Stats.ELV \ 10, 1) Then
                            Exit Function
                        End If
                
                    ' Con galeón no pierde los últimos 6 * (cada 10 niveles; max 3) slots
110                 ElseIf ObjData(.Invent.BarcoObjIndex).Ropaje = iGaleon Then
                
112                     If slot > .CurrentInventorySlots - 6 * min(.Stats.ELV \ 10, 3) Then
                            Exit Function
                        End If
                
                    End If
                
                End If
            
            End If
        
        End With
    
114     PirataCaeItem = True

        
        Exit Function

PirataCaeItem_Err:
116     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.PirataCaeItem", Erl)

        
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
        
        On Error GoTo TirarTodosLosItems_Err

        Dim i         As Byte
        Dim NuevaPos  As WorldPos
        Dim MiObj     As obj
        Dim ItemIndex As Integer
    
100     With UserList(UserIndex)
    
102         For i = 1 To .CurrentInventorySlots
    
104             ItemIndex = .Invent.Object(i).ObjIndex

106             If ItemIndex > 0 Then

108                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
110                     NuevaPos.X = 0
112                     NuevaPos.Y = 0
                    
114                     MiObj.Amount = .Invent.Object(i).Amount
116                     MiObj.ObjIndex = ItemIndex

                        If .flags.CarroMineria = 1 Then
118                         If ItemIndex = ORO_MINA Or ItemIndex = PLATA_MINA Or ItemIndex = HIERRO_MINA Then
120                             MiObj.Amount = MiObj.Amount * 0.3
                            End If
                        End If
                    
122                     Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
            
124                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
126                         Call DropObj(UserIndex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        
                        ' WyroX: Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                        Else
128                         Call QuitarUserInvItem(UserIndex, i, MiObj.Amount)
130                         Call UpdateUserInv(False, UserIndex, i)
                        End If
                
                    End If

                End If
    
138         Next i
    
        End With
 
        Exit Sub

TirarTodosLosItems_Err:
140     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.TirarTodosLosItems", Erl)

142     Resume Next
        
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo ItemNewbie_Err
        

100     ItemNewbie = ObjData(ItemIndex).Newbie = 1

        
        Exit Function

ItemNewbie_Err:
102     Call RegistrarError(Err.Number, Err.Description, "InvUsuario.ItemNewbie", Erl)
104     Resume Next
        
End Function
