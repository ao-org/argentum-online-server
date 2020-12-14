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

Public Function TieneObjetosRobables(ByVal Userindex As Integer) As Boolean

        '17/09/02
        'Agregue que la función se asegure que el objeto no es un barco

        On Error Resume Next

        Dim i        As Integer

        Dim ObjIndex As Integer

100     For i = 1 To UserList(Userindex).CurrentInventorySlots
102         ObjIndex = UserList(Userindex).Invent.Object(i).ObjIndex

104         If ObjIndex > 0 Then
106             If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And ObjData(ObjIndex).OBJType <> eOBJType.OtDonador And ObjData(ObjIndex).OBJType <> eOBJType.otRunas) Then
108                 TieneObjetosRobables = True
                    Exit Function

                End If
    
            End If

110     Next i

End Function

Function ClasePuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer, Optional slot As Byte) As Boolean

        On Error GoTo manejador

        'Call LogTarea("ClasePuedeUsarItem")

        Dim flag As Boolean

100     If slot <> 0 Then
102         If UserList(Userindex).Invent.Object(slot).Equipped Then
104             ClasePuedeUsarItem = True
                Exit Function

            End If

        End If

106     If EsGM(Userindex) Then
108         ClasePuedeUsarItem = True
            Exit Function

        End If

        'Admins can use ANYTHING!
        'If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        'If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer

110     For i = 1 To NUMCLASES

112         If ObjData(ObjIndex).ClaseProhibida(i) = UserList(Userindex).clase Then
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

Sub QuitarNewbieObj(ByVal Userindex As Integer)
        
        On Error GoTo QuitarNewbieObj_Err
        

        Dim j As Integer

100     For j = 1 To UserList(Userindex).CurrentInventorySlots

102         If UserList(Userindex).Invent.Object(j).ObjIndex > 0 Then
             
104             If ObjData(UserList(Userindex).Invent.Object(j).ObjIndex).Newbie = 1 Then
106                 Call QuitarUserInvItem(Userindex, j, MAX_INVENTORY_OBJS)
108                 Call UpdateUserInv(False, Userindex, j)

                End If
        
            End If

110     Next j
    
        'Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
        'es transportado a su hogar de origen ;)
112     If UCase$(MapInfo(UserList(Userindex).Pos.Map).restrict_mode) = "NEWBIE" Then
        
            Dim DeDonde As WorldPos
        
114         Select Case UserList(Userindex).Hogar

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
        
142         Call WarpUserChar(Userindex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    
        End If

        
        Exit Sub

QuitarNewbieObj_Err:
144     Call RegistrarError(Err.Number, Err.description, "InvUsuario.QuitarNewbieObj", Erl)
146     Resume Next
        
End Sub

Sub LimpiarInventario(ByVal Userindex As Integer)
        
        On Error GoTo LimpiarInventario_Err
        

        Dim j As Integer

100     For j = 1 To UserList(Userindex).CurrentInventorySlots
102         UserList(Userindex).Invent.Object(j).ObjIndex = 0
104         UserList(Userindex).Invent.Object(j).Amount = 0
106         UserList(Userindex).Invent.Object(j).Equipped = 0
        
        Next

108     UserList(Userindex).Invent.NroItems = 0

110     UserList(Userindex).Invent.ArmourEqpObjIndex = 0
112     UserList(Userindex).Invent.ArmourEqpSlot = 0

114     UserList(Userindex).Invent.WeaponEqpObjIndex = 0
116     UserList(Userindex).Invent.WeaponEqpSlot = 0

118     UserList(Userindex).Invent.HerramientaEqpObjIndex = 0
120     UserList(Userindex).Invent.HerramientaEqpSlot = 0

122     UserList(Userindex).Invent.CascoEqpObjIndex = 0
124     UserList(Userindex).Invent.CascoEqpSlot = 0

126     UserList(Userindex).Invent.EscudoEqpObjIndex = 0
128     UserList(Userindex).Invent.EscudoEqpSlot = 0

130     UserList(Userindex).Invent.AnilloEqpObjIndex = 0
132     UserList(Userindex).Invent.AnilloEqpSlot = 0

134     UserList(Userindex).Invent.NudilloObjIndex = 0
136     UserList(Userindex).Invent.NudilloSlot = 0

138     UserList(Userindex).Invent.MunicionEqpObjIndex = 0
140     UserList(Userindex).Invent.MunicionEqpSlot = 0

142     UserList(Userindex).Invent.BarcoObjIndex = 0
144     UserList(Userindex).Invent.BarcoSlot = 0

146     UserList(Userindex).Invent.MonturaObjIndex = 0
148     UserList(Userindex).Invent.MonturaSlot = 0

150     UserList(Userindex).Invent.MagicoObjIndex = 0
152     UserList(Userindex).Invent.MagicoSlot = 0

        
        Exit Sub

LimpiarInventario_Err:
154     Call RegistrarError(Err.Number, Err.description, "InvUsuario.LimpiarInventario", Erl)
156     Resume Next
        
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal Userindex As Integer)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
        '***************************************************
        On Error GoTo ErrHandler

        'If Cantidad > 100000 Then Exit Sub
100     If UserList(Userindex).flags.BattleModo = 1 Then Exit Sub

        'SI EL Pjta TIENE ORO LO TIRAMOS
102     If (Cantidad > 0) And (Cantidad <= UserList(Userindex).Stats.GLD) Then

            Dim i     As Byte

            Dim MiObj As obj

            Dim Logs  As Long

            'info debug
            Dim loops As Integer
        
104         Logs = Cantidad

            Dim Extra    As Long

            Dim TeniaOro As Long

106         TeniaOro = UserList(Userindex).Stats.GLD

108         If Cantidad > 500000 Then 'Para evitar explotar demasiado
110             Extra = Cantidad - 500000
112             Cantidad = 500000

            End If
        
114         Do While (Cantidad > 0)
            
116             If Cantidad > MAX_INVENTORY_OBJS And UserList(Userindex).Stats.GLD > MAX_INVENTORY_OBJS Then
118                 MiObj.Amount = MAX_INVENTORY_OBJS
120                 Cantidad = Cantidad - MiObj.Amount
                Else
122                 MiObj.Amount = Cantidad
124                 Cantidad = Cantidad - MiObj.Amount

                End If

126             MiObj.ObjIndex = iORO

                Dim AuxPos As WorldPos
128             If UserList(Userindex).clase = eClass.Pirat Then
130                 AuxPos = TirarItemAlPiso(UserList(Userindex).Pos, MiObj, False)
                Else
132                 AuxPos = TirarItemAlPiso(UserList(Userindex).Pos, MiObj, True)
                End If
            
134             If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
136                 UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - MiObj.Amount

                End If
            
                'info debug
138             loops = loops + 1

140             If loops > 100 Then
142                 LogError ("Error en tiraroro")
                    Exit Sub

                End If
            
            Loop
        
144         If EsGM(Userindex) Then
146             If MiObj.ObjIndex = iORO Then
148                 Call LogGM(UserList(Userindex).name, "Tiro: " & Logs & " monedas de oro.")
                Else
150                 Call LogGM(UserList(Userindex).name, "Tiro cantidad:" & Logs & " Objeto:" & ObjData(MiObj.ObjIndex).name)

                End If

            End If
        
152         If TeniaOro = UserList(Userindex).Stats.GLD Then Extra = 0
154         If Extra > 0 Then
156             UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - Extra

            End If
    
        End If

        Exit Sub

ErrHandler:

End Sub

Sub QuitarUserInvItem(ByVal Userindex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarUserInvItem_Err
        

100     If slot < 1 Or slot > UserList(Userindex).CurrentInventorySlots Then Exit Sub
    
102     With UserList(Userindex).Invent.Object(slot)

104         If .Amount <= Cantidad And .Equipped = 1 Then
106             Call Desequipar(Userindex, slot)

            End If
        
            'Quita un objeto
108         .Amount = .Amount - Cantidad

            '¿Quedan mas?
110         If .Amount <= 0 Then
112             UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
114             .ObjIndex = 0
116             .Amount = 0

            End If

        End With

        
        Exit Sub

QuitarUserInvItem_Err:
118     Call RegistrarError(Err.Number, Err.description, "InvUsuario.QuitarUserInvItem", Erl)
120     Resume Next
        
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateUserInv_Err
        

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(Userindex).Invent.Object(slot).ObjIndex > 0 Then
104             Call ChangeUserInv(Userindex, slot, UserList(Userindex).Invent.Object(slot))
            Else
106             Call ChangeUserInv(Userindex, slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To UserList(Userindex).CurrentInventorySlots

                'Actualiza el inventario
110             If UserList(Userindex).Invent.Object(LoopC).ObjIndex > 0 Then
112                 Call ChangeUserInv(Userindex, LoopC, UserList(Userindex).Invent.Object(LoopC))
                Else
114                 Call ChangeUserInv(Userindex, LoopC, NullObj)

                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateUserInv_Err:
118     Call RegistrarError(Err.Number, Err.description, "InvUsuario.UpdateUserInv", Erl)
120     Resume Next
        
End Sub

Sub DropObj(ByVal Userindex As Integer, _
            ByVal slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        
        On Error GoTo DropObj_Err

        Dim obj As obj

100     If num > 0 Then
            
102         With UserList(Userindex)

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
                        
124                     Call QuitarUserInvItem(Userindex, slot, num)
126                     Call UpdateUserInv(False, Userindex, slot)
                        
128                     If Not .flags.Privilegios And PlayerType.user Then
130                         Call LogGM(.name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)
                        End If
    
                    Else
                    
                        'Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
132                     Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                Else
134                 Call QuitarUserInvItem(Userindex, slot, num)
136                 Call UpdateUserInv(False, Userindex, slot)
    
                End If
            
            End With

        End If
        
        Exit Sub

DropObj_Err:
138     Call RegistrarError(Err.Number, Err.description, "InvUsuario.DropObj", Erl)

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
114     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EraseObj", Erl)
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
108                 Call AgregarItemLimpiza(Map, X, Y, MapData(Map, X, Y).ObjInfo.ObjIndex <> 0)
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
120     Call RegistrarError(Err.Number, Err.description, "InvUsuario.MakeObj", Erl)

122     Resume Next
        
End Sub

Function MeterItemEnInventario(ByVal Userindex As Integer, ByRef MiObj As obj) As Boolean

        On Error GoTo ErrHandler

        'Call LogTarea("MeterItemEnInventario")
 
        Dim X    As Integer

        Dim Y    As Integer

        Dim slot As Byte

        '¿el user ya tiene un objeto del mismo tipo? ?????
100     If MiObj.ObjIndex = 12 Then
102         UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + MiObj.Amount

        Else
    
104         slot = 1

106         Do Until UserList(Userindex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And UserList(Userindex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
108             slot = slot + 1

110             If slot > UserList(Userindex).CurrentInventorySlots Then
                    Exit Do

                End If

            Loop
        
            'Sino busca un slot vacio
112         If slot > UserList(Userindex).CurrentInventorySlots Then
114             slot = 1

116             Do Until UserList(Userindex).Invent.Object(slot).ObjIndex = 0
118                 slot = slot + 1

120                 If slot > UserList(Userindex).CurrentInventorySlots Then
                        'Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
122                     Call WriteLocaleMsg(Userindex, "328", FontTypeNames.FONTTYPE_FIGHT)
124                     MeterItemEnInventario = False
                        Exit Function

                    End If

                Loop
126             UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems + 1

            End If
        
            'Mete el objeto
128         If UserList(Userindex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
                'Menor que MAX_INV_OBJS
130             UserList(Userindex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex
132             UserList(Userindex).Invent.Object(slot).Amount = UserList(Userindex).Invent.Object(slot).Amount + MiObj.Amount
            Else
134             UserList(Userindex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS

            End If
        
136         MeterItemEnInventario = True
           
138         Call UpdateUserInv(False, Userindex, slot)

        End If

140     WriteUpdateGold (Userindex)
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

Sub GetObj(ByVal Userindex As Integer)
        
        On Error GoTo GetObj_Err
        

        Dim obj   As ObjData

        Dim MiObj As obj

        '¿Hay algun obj?
100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
102         If ObjData(MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

                Dim X    As Integer

                Dim Y    As Integer

                Dim slot As Byte
        
104             X = UserList(Userindex).Pos.X
106             Y = UserList(Userindex).Pos.Y
108             obj = ObjData(MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).ObjInfo.ObjIndex)
110             MiObj.Amount = MapData(UserList(Userindex).Pos.Map, X, Y).ObjInfo.Amount
112             MiObj.ObjIndex = MapData(UserList(Userindex).Pos.Map, X, Y).ObjInfo.ObjIndex
        
114             If Not MeterItemEnInventario(Userindex, MiObj) Then
                    'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
                Else
            
                    'Quitamos el objeto
116                 Call EraseObj(MapData(UserList(Userindex).Pos.Map, X, Y).ObjInfo.Amount, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y)

118                 If Not UserList(Userindex).flags.Privilegios And PlayerType.user Then Call LogGM(UserList(Userindex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
    
120                 If BusquedaTesoroActiva Then
122                     If UserList(Userindex).Pos.Map = TesoroNumMapa And UserList(Userindex).Pos.X = TesoroX And UserList(Userindex).Pos.Y = TesoroY Then
    
124                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(Userindex).name & " encontro el tesoro ¡Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
126                         BusquedaTesoroActiva = False

                        End If

                    End If
                
128                 If BusquedaRegaloActiva Then
130                     If UserList(Userindex).Pos.Map = RegaloNumMapa And UserList(Userindex).Pos.X = RegaloX And UserList(Userindex).Pos.Y = RegaloY Then
132                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(Userindex).name & " fue el valiente que encontro el gran item magico ¡Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
134                         BusquedaRegaloActiva = False

                        End If

                    End If
                
                    'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                    'Es un Objeto que tenemos que loguear?
136                 If ObjData(MiObj.ObjIndex).Log = 1 Then
138                     Call LogDesarrollo(UserList(Userindex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)

                        ' ElseIf MiObj.Amount = 1000 Then 'Es mucha cantidad?
                        '  'Si no es de los prohibidos de loguear, lo logueamos.
                        '   'If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                        ' Call LogDesarrollo(UserList(UserIndex).name & " juntó del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
                        ' End If
                    End If
                
                End If

            End If

        Else

140         If Not UserList(Userindex).flags.UltimoMensaje = 261 Then
142             Call WriteLocaleMsg(Userindex, "261", FontTypeNames.FONTTYPE_INFO)
144             UserList(Userindex).flags.UltimoMensaje = 261

            End If
    
            'Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
        End If

        
        Exit Sub

GetObj_Err:
146     Call RegistrarError(Err.Number, Err.description, "InvUsuario.GetObj", Erl)
148     Resume Next
        
End Sub

Sub Desequipar(ByVal Userindex As Integer, ByVal slot As Byte)
        
        On Error GoTo Desequipar_Err

        'Desequipa el item slot del inventario
        Dim obj As ObjData

100     If (slot < LBound(UserList(Userindex).Invent.Object)) Or (slot > UBound(UserList(Userindex).Invent.Object)) Then
            Exit Sub
102     ElseIf UserList(Userindex).Invent.Object(slot).ObjIndex = 0 Then
            Exit Sub

        End If

104     obj = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex)

106     Select Case obj.OBJType

            Case eOBJType.otWeapon
108             UserList(Userindex).Invent.Object(slot).Equipped = 0
110             UserList(Userindex).Invent.WeaponEqpObjIndex = 0
112             UserList(Userindex).Invent.WeaponEqpSlot = 0
114             UserList(Userindex).Char.Arma_Aura = ""
116             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 1))
        
118             UserList(Userindex).Char.WeaponAnim = NingunArma
            
120             If UserList(Userindex).flags.Montado = 0 Then
122                 Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                End If
                
124             If obj.MagicDamageBonus > 0 Then
126                 Call WriteUpdateDM(Userindex)
                End If
    
128         Case eOBJType.otFlechas
130             UserList(Userindex).Invent.Object(slot).Equipped = 0
132             UserList(Userindex).Invent.MunicionEqpObjIndex = 0
134             UserList(Userindex).Invent.MunicionEqpSlot = 0
    
                ' Case eOBJType.otAnillos
                '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
                '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
                ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
            
136         Case eOBJType.otHerramientas
138             UserList(Userindex).Invent.Object(slot).Equipped = 0
140             UserList(Userindex).Invent.HerramientaEqpObjIndex = 0
142             UserList(Userindex).Invent.HerramientaEqpSlot = 0

144             If UserList(Userindex).flags.UsandoMacro = True Then
146                 Call WriteMacroTrabajoToggle(Userindex, False)
                End If
        
148             UserList(Userindex).Char.WeaponAnim = NingunArma
            
150             If UserList(Userindex).flags.Montado = 0 Then
152                 Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                End If
       
154         Case eOBJType.otmagicos
    
156             Select Case obj.EfectoMagico

                    Case 1 'Regenera Energia
158                     UserList(Userindex).flags.RegeneracionSta = 0

160                 Case 2 'Modifica los Atributos
162                     UserList(Userindex).Stats.UserAtributos(obj.QueAtributo) = UserList(Userindex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                
164                     UserList(Userindex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(Userindex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                        ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
166                     Call WriteFYA(Userindex)

168                 Case 3 'Modifica los skills
170                     UserList(Userindex).Stats.UserSkills(obj.QueSkill) = UserList(Userindex).Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento

172                 Case 4 ' Regeneracion Vida
174                     UserList(Userindex).flags.RegeneracionHP = 0

176                 Case 5 ' Regeneracion Mana
178                     UserList(Userindex).flags.RegeneracionMana = 0

180                 Case 6 'Aumento Golpe
182                     UserList(Userindex).Stats.MaxHit = UserList(Userindex).Stats.MaxHit - obj.CuantoAumento
184                     UserList(Userindex).Stats.MinHIT = UserList(Userindex).Stats.MinHIT - obj.CuantoAumento

186                 Case 7 '
                
188                 Case 9 ' Orbe Ignea
190                     UserList(Userindex).flags.NoMagiaEfeceto = 0

192                 Case 10
194                     UserList(Userindex).flags.incinera = 0

196                 Case 11
198                     UserList(Userindex).flags.Paraliza = 0

200                 Case 12

202                     If UserList(Userindex).flags.Muerto = 0 Then UserList(Userindex).flags.CarroMineria = 0
                
204                 Case 14
                        'UserList(UserIndex).flags.DañoMagico = 0
                
206                 Case 15 'Pendiete del Sacrificio
208                     UserList(Userindex).flags.PendienteDelSacrificio = 0
                 
210                 Case 16
212                     UserList(Userindex).flags.NoPalabrasMagicas = 0

214                 Case 17 'Sortija de la verdad
216                     UserList(Userindex).flags.NoDetectable = 0

218                 Case 18 ' Pendiente del Experto
220                     UserList(Userindex).flags.PendienteDelExperto = 0

222                 Case 19
224                     UserList(Userindex).flags.Envenena = 0

226                 Case 20 ' anillo de las sombras
228                     UserList(Userindex).flags.AnilloOcultismo = 0
                
                End Select
        
230             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 5))
232             UserList(Userindex).Char.Otra_Aura = 0
234             UserList(Userindex).Invent.Object(slot).Equipped = 0
236             UserList(Userindex).Invent.MagicoObjIndex = 0
238             UserList(Userindex).Invent.MagicoSlot = 0
        
240         Case eOBJType.otNUDILLOS
    
                'falta mandar animacion
            
242             UserList(Userindex).Invent.Object(slot).Equipped = 0
244             UserList(Userindex).Invent.NudilloObjIndex = 0
246             UserList(Userindex).Invent.NudilloSlot = 0
        
248             UserList(Userindex).Char.Arma_Aura = ""
250             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 1))
        
252             UserList(Userindex).Char.WeaponAnim = NingunArma
254             Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        
256         Case eOBJType.otArmadura
258             UserList(Userindex).Invent.Object(slot).Equipped = 0
260             UserList(Userindex).Invent.ArmourEqpObjIndex = 0
262             UserList(Userindex).Invent.ArmourEqpSlot = 0
        
264             If UserList(Userindex).flags.Navegando = 0 Then
266                 If UserList(Userindex).flags.Montado = 0 Then
268                     Call DarCuerpoDesnudo(Userindex)
270                     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                    End If
                End If
        
272             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 2))
        
274             UserList(Userindex).Char.Body_Aura = 0

276             If obj.ResistenciaMagica > 0 Then
278                 Call WriteUpdateRM(Userindex)
                End If
    
280         Case eOBJType.otCASCO
282             UserList(Userindex).Invent.Object(slot).Equipped = 0
284             UserList(Userindex).Invent.CascoEqpObjIndex = 0
286             UserList(Userindex).Invent.CascoEqpSlot = 0
288             UserList(Userindex).Char.Head_Aura = 0
290             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 4))

292             UserList(Userindex).Char.CascoAnim = NingunCasco
294             Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
    
296             If obj.ResistenciaMagica > 0 Then
298                 Call WriteUpdateRM(Userindex)
                End If
    
300         Case eOBJType.otESCUDO
302             UserList(Userindex).Invent.Object(slot).Equipped = 0
304             UserList(Userindex).Invent.EscudoEqpObjIndex = 0
306             UserList(Userindex).Invent.EscudoEqpSlot = 0
308             UserList(Userindex).Char.Escudo_Aura = 0
310             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 3))
        
312             UserList(Userindex).Char.ShieldAnim = NingunEscudo

314             If UserList(Userindex).flags.Montado = 0 Then
316                 Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                End If
                
318             If obj.ResistenciaMagica > 0 Then
320                 Call WriteUpdateRM(Userindex)
                End If
                
322         Case eOBJType.otAnillos
324             UserList(Userindex).Invent.Object(slot).Equipped = 0
326             UserList(Userindex).Invent.AnilloEqpObjIndex = 0
328             UserList(Userindex).Invent.AnilloEqpSlot = 0
330             UserList(Userindex).Char.Anillo_Aura = 0
332             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 6))

334             If obj.MagicDamageBonus > 0 Then
336                 Call WriteUpdateDM(Userindex)
                End If
                
338             If obj.ResistenciaMagica > 0 Then
340                 Call WriteUpdateRM(Userindex)
                End If
        
        End Select

342     Call UpdateUserInv(False, Userindex, slot)

        
        Exit Sub

Desequipar_Err:
344     Call RegistrarError(Err.Number, Err.description, "InvUsuario.Desequipar", Erl)
346     Resume Next
        
End Sub

Function SexoPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(Userindex) Then
102         SexoPuedeUsarItem = True
            Exit Function

        End If

104     If ObjData(ObjIndex).Mujer = 1 Then
106         SexoPuedeUsarItem = UserList(Userindex).genero <> eGenero.Hombre
108     ElseIf ObjData(ObjIndex).Hombre = 1 Then
110         SexoPuedeUsarItem = UserList(Userindex).genero <> eGenero.Mujer
        Else
112         SexoPuedeUsarItem = True

        End If

        Exit Function
ErrHandler:
114     Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean
        
        On Error GoTo FaccionPuedeUsarItem_Err
        

100     If ObjData(ObjIndex).Real = 1 Then
102         If Status(Userindex) = 3 Then
104             FaccionPuedeUsarItem = esArmada(Userindex)
            Else
106             FaccionPuedeUsarItem = False

            End If

108     ElseIf ObjData(ObjIndex).Caos = 1 Then

110         If Status(Userindex) = 2 Then
112             FaccionPuedeUsarItem = esCaos(Userindex)
            Else
114             FaccionPuedeUsarItem = False

            End If

        Else
116         FaccionPuedeUsarItem = True

        End If

        
        Exit Function

FaccionPuedeUsarItem_Err:
118     Call RegistrarError(Err.Number, Err.description, "InvUsuario.FaccionPuedeUsarItem", Erl)
120     Resume Next
        
End Function

Sub EquiparInvItem(ByVal Userindex As Integer, ByVal slot As Byte)

        On Error GoTo ErrHandler

        Dim errordesc As String

        'Equipa un item del inventario
        Dim obj       As ObjData
        Dim ObjIndex  As Integer

100     ObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
102     obj = ObjData(ObjIndex)

104     If obj.Newbie = 1 And Not EsNewbie(Userindex) And Not EsGM(Userindex) Then
106         Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

108     If UserList(Userindex).Stats.ELV < obj.MinELV And Not EsGM(Userindex) Then
110         Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
112     If obj.SkillIndex > 0 Then
    
114         If UserList(Userindex).Stats.UserSkills(obj.SkillIndex) < obj.SkillRequerido And Not EsGM(Userindex) Then
116             Call WriteConsoleMsg(Userindex, "Necesitas " & obj.SkillRequerido & " puntos en " & SkillsNames(obj.SkillIndex) & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

        End If
    
118     With UserList(Userindex)
    
120         Select Case obj.OBJType

                Case eOBJType.otWeapon
                
122                 errordesc = "Arma"

124                 If Not ClasePuedeUsarItem(Userindex, ObjIndex, slot) And FaccionPuedeUsarItem(Userindex, ObjIndex) Then
126                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
128                 If Not FaccionPuedeUsarItem(Userindex, ObjIndex) Then
130                     Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Si esta equipado lo quita
132                 If .Invent.Object(slot).Equipped Then
                    
                        'Quitamos del inv el item
134                     Call Desequipar(Userindex, slot)
                        
                        'Animacion por defecto
136                     .Char.WeaponAnim = NingunArma

138                     If .flags.Montado = 0 Then
140                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If

                        Exit Sub

                    End If
            
                    'Quitamos el elemento anterior
142                 If .Invent.WeaponEqpObjIndex > 0 Then
144                     Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
                    End If
            
146                 If .Invent.HerramientaEqpObjIndex > 0 Then
148                     Call Desequipar(Userindex, .Invent.HerramientaEqpSlot)
                    End If
            
150                 If .Invent.NudilloObjIndex > 0 Then
152                     Call Desequipar(Userindex, .Invent.NudilloSlot)
                    End If
            
154                 .Invent.Object(slot).Equipped = 1
156                 .Invent.WeaponEqpObjIndex = .Invent.Object(slot).ObjIndex
158                 .Invent.WeaponEqpSlot = slot
            
160                 If obj.proyectil = 1 Then 'Si es un arco, desequipa el escudo.
            
                        'If .Invent.EscudoEqpObjIndex = 404 Or .Invent.EscudoEqpObjIndex = 1007 Or .Invent.EscudoEqpObjIndex = 1358 Then
162                     If .Invent.EscudoEqpObjIndex = 1700 Or _
                           .Invent.EscudoEqpObjIndex = 1730 Or _
                           .Invent.EscudoEqpObjIndex = 1724 Or _
                           .Invent.EscudoEqpObjIndex = 1717 Or _
                           .Invent.EscudoEqpObjIndex = 1699 Then
                
                        Else

164                         If .Invent.EscudoEqpObjIndex > 0 Then
166                             Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
168                             Call WriteConsoleMsg(Userindex, "No podes tirar flechas si tenés un escudo equipado. Tu escudo fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        End If

                    End If
            
                    'Sonido
170                 If obj.SndAura = 0 Then
172                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
174                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
            
176                 If Len(obj.CreaGRH) <> 0 Then
178                     .Char.Arma_Aura = obj.CreaGRH
180                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If
                
182                 If obj.MagicDamageBonus > 0 Then
184                     Call WriteUpdateDM(Userindex)
                    End If
                
186                 If .flags.Montado = 0 Then
                
188                     If .flags.Navegando = 0 Then
190                         .Char.WeaponAnim = obj.WeaponAnim
192                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
      
194             Case eOBJType.otHerramientas
        
196                 If Not ClasePuedeUsarItem(Userindex, ObjIndex, slot) Then
198                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
200                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
202                     Call Desequipar(Userindex, slot)
                        Exit Sub

                    End If

204                 If obj.MinSkill <> 0 Then
                
206                     If .Stats.UserSkills(obj.QueSkill) < obj.MinSkill Then
208                         Call WriteConsoleMsg(Userindex, "Para podes usar " & obj.name & " necesitas al menos " & obj.MinSkill & " puntos en " & SkillsNames(obj.QueSkill) & ".", FontTypeNames.FONTTYPE_INFOIAO)
                            Exit Sub
                        End If

                    End If

                    'Quitamos el elemento anterior
210                 If .Invent.HerramientaEqpObjIndex > 0 Then
212                     Call Desequipar(Userindex, .Invent.HerramientaEqpSlot)
                    End If
             
214                 If .Invent.WeaponEqpObjIndex > 0 Then
216                     Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
                    End If
             
218                 .Invent.Object(slot).Equipped = 1
220                 .Invent.HerramientaEqpObjIndex = ObjIndex
222                 .Invent.HerramientaEqpSlot = slot
             
224                 If .flags.Montado = 0 Then
                
226                     If .flags.Navegando = 0 Then
228                         .Char.WeaponAnim = obj.WeaponAnim
230                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
       
232             Case eOBJType.otmagicos
            
234                 errordesc = "Magico"
    
236                 If .flags.Muerto = 1 Then
238                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
        
                    'Si esta equipado lo quita
240                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
242                     Call Desequipar(Userindex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
244                 If .Invent.MagicoObjIndex > 0 Then
246                     Call Desequipar(Userindex, .Invent.MagicoSlot)
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
                        
262                         .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento
                        
264                         If .Stats.UserAtributos(obj.QueAtributo) > MAXATRIBUTOS Then
266                             .Stats.UserAtributos(obj.QueAtributo) = MAXATRIBUTOS
                            End If
                
268                         Call WriteFYA(Userindex)

270                     Case 3 'Modifica los skills
            
272                         .Stats.UserSkills(obj.QueSkill) = .Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento

274                     Case 4
276                         .flags.RegeneracionHP = 1

278                     Case 5
280                         .flags.RegeneracionMana = 1

282                     Case 6
                            'Call WriteConsoleMsg(UserIndex, "Item, temporalmente deshabilitado.", FontTypeNames.FONTTYPE_INFO)
284                         .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
286                         .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento

288                     Case 9
290                         .flags.NoMagiaEfeceto = 1

292                     Case 10
294                         .flags.incinera = 1

296                     Case 11
298                         .flags.Paraliza = 1

300                     Case 12
302                         .flags.CarroMineria = 1
                
304                     Case 14
                            '.flags.DañoMagico = obj.CuantoAumento
                
306                     Case 15 'Pendiete del Sacrificio
308                         .flags.PendienteDelSacrificio = 1

310                     Case 16
312                         .flags.NoPalabrasMagicas = 1

314                     Case 17
316                         .flags.NoDetectable = 1
                   
318                     Case 18 ' Pendiente del Experto
320                         .flags.PendienteDelExperto = 1

322                     Case 19
324                         .flags.Envenena = 1

326                     Case 20 'Anillo ocultismo
328                         .flags.AnilloOcultismo = 1
    
                    End Select
            
                    'Sonido
330                 If obj.SndAura <> 0 Then
332                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
            
334                 If Len(obj.CreaGRH) <> 0 Then
336                     .Char.Otra_Aura = obj.CreaGRH
338                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
                    End If
        
                    'Call WriteUpdateExp(UserIndex)
                    'Call CheckUserLevel(UserIndex)
            
340             Case eOBJType.otNUDILLOS
    
342                 If .flags.Muerto = 1 Then
344                     Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
346                 If Not ClasePuedeUsarItem(Userindex, ObjIndex, slot) Then
348                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                 
350                 If .Invent.WeaponEqpObjIndex > 0 Then
352                     Call Desequipar(Userindex, .Invent.WeaponEqpSlot)

                    End If

354                 If .Invent.Object(slot).Equipped Then
356                     Call Desequipar(Userindex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
358                 If .Invent.NudilloObjIndex > 0 Then
360                     Call Desequipar(Userindex, .Invent.NudilloSlot)

                    End If
        
362                 .Invent.Object(slot).Equipped = 1
364                 .Invent.NudilloObjIndex = .Invent.Object(slot).ObjIndex
366                 .Invent.NudilloSlot = slot
        
                    'Falta enviar anim
368                 If .flags.Montado = 0 Then
                
370                     If .flags.Navegando = 0 Then
372                         .Char.WeaponAnim = obj.WeaponAnim
374                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
            
376                 If obj.SndAura = 0 Then
378                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
380                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
                 
382                 If Len(obj.CreaGRH) <> 0 Then
384                     .Char.Arma_Aura = obj.CreaGRH
386                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If
    
388             Case eOBJType.otFlechas

390                 If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Or Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
392                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
394                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
396                     Call Desequipar(Userindex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
398                 If .Invent.MunicionEqpObjIndex > 0 Then
400                     Call Desequipar(Userindex, .Invent.MunicionEqpSlot)
                    End If
        
402                 .Invent.Object(slot).Equipped = 1
404                 .Invent.MunicionEqpObjIndex = .Invent.Object(slot).ObjIndex
406                 .Invent.MunicionEqpSlot = slot

408             Case eOBJType.otArmadura
                
410                 If obj.Ropaje = 0 Then
412                     Call WriteConsoleMsg(Userindex, "Hay un error con este objeto. Infórmale a un administrador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Nos aseguramos que puede usarla
414                 If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Or _
                       Not SexoPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Or _
                       Not CheckRazaUsaRopa(Userindex, .Invent.Object(slot).ObjIndex) Or _
                       Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
                    
416                     Call WriteConsoleMsg(Userindex, "Tu clase, género, raza o facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
418                 If .Invent.Object(slot).Equipped Then
                    
420                     Call Desequipar(Userindex, slot)

422                     If .flags.Navegando = 0 Then
                        
424                         If .flags.Montado = 0 Then
426                             Call DarCuerpoDesnudo(Userindex)
428                             Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                            End If

                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
430                 If .Invent.ArmourEqpObjIndex > 0 Then
432                     errordesc = "Armadura 2"
434                     Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
436                     errordesc = "Armadura 3"

                    End If
  
                    'Lo equipa
438                 If Len(obj.CreaGRH) <> 0 Then
440                     .Char.Body_Aura = obj.CreaGRH
442                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))

                    End If
            
444                 .Invent.Object(slot).Equipped = 1
446                 .Invent.ArmourEqpObjIndex = .Invent.Object(slot).ObjIndex
448                 .Invent.ArmourEqpSlot = slot
                            
450                 If .flags.Montado = 0 Then
                
452                     If .flags.Navegando = 0 Then
                        
454                         .Char.Body = obj.Ropaje
                
456                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
458                         .flags.Desnudo = 0
            
                        End If

                    End If
                
460                 If obj.ResistenciaMagica > 0 Then
462                     Call WriteUpdateRM(Userindex)
                    End If
    
464             Case eOBJType.otCASCO
                
466                 If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Then
468                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
470                 If Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
472                     Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    
                    End If
                
                    'Si esta equipado lo quita
474                 If .Invent.Object(slot).Equipped Then
476                     Call Desequipar(Userindex, slot)
                
478                     .Char.CascoAnim = NingunCasco
480                     Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Exit Sub

                    End If
    
                    'Quita el anterior
482                 If .Invent.CascoEqpObjIndex > 0 Then
484                     Call Desequipar(Userindex, .Invent.CascoEqpSlot)
                    End If
            
486                 errordesc = "Casco"

                    'Lo equipa
488                 If Len(obj.CreaGRH) <> 0 Then
490                     .Char.Head_Aura = obj.CreaGRH
492                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
                    End If
            
494                 .Invent.Object(slot).Equipped = 1
496                 .Invent.CascoEqpObjIndex = .Invent.Object(slot).ObjIndex
498                 .Invent.CascoEqpSlot = slot
            
500                 If .flags.Navegando = 0 Then
502                     .Char.CascoAnim = obj.CascoAnim
504                     Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                
506                 If obj.ResistenciaMagica > 0 Then
508                     Call WriteUpdateRM(Userindex)
                    End If

510             Case eOBJType.otESCUDO

512                 If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Then
514                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
516                 If Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
518                     Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
520                 If .Invent.Object(slot).Equipped Then
522                     Call Desequipar(Userindex, slot)
                 
524                     .Char.ShieldAnim = NingunEscudo

526                     If .flags.Montado = 0 Then
528                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
     
                    'Quita el anterior
530                 If .Invent.EscudoEqpObjIndex > 0 Then
532                     Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
                    End If
     
                    'Lo equipa
             
534                 If .Invent.Object(slot).ObjIndex = 1700 Or _
                       .Invent.Object(slot).ObjIndex = 1730 Or _
                       .Invent.Object(slot).ObjIndex = 1724 Or _
                       .Invent.Object(slot).ObjIndex = 1717 Or _
                       .Invent.Object(slot).ObjIndex = 1699 Then
             
                    Else

536                     If .Invent.WeaponEqpObjIndex > 0 Then
538                         If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
540                             Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
542                             Call WriteConsoleMsg(Userindex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                            End If
                        End If

                    End If
            
544                 errordesc = "Escudo"
             
546                 If Len(obj.CreaGRH) <> 0 Then
548                     .Char.Escudo_Aura = obj.CreaGRH
550                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
                    End If

552                 .Invent.Object(slot).Equipped = 1
554                 .Invent.EscudoEqpObjIndex = .Invent.Object(slot).ObjIndex
556                 .Invent.EscudoEqpSlot = slot
                 
558                 If .flags.Navegando = 0 Then
560                     If .flags.Montado = 0 Then
562                         .Char.ShieldAnim = obj.ShieldAnim
564                         Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                    End If
                
566                 If obj.ResistenciaMagica > 0 Then
568                     Call WriteUpdateRM(Userindex)
                    End If
                
570             Case eOBJType.otAnillos

572                 If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Then
574                     Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

576                 If Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
578                     Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
580                 If .Invent.Object(slot).Equipped Then
582                     Call Desequipar(Userindex, slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
584                 If .Invent.AnilloEqpSlot > 0 Then
586                     Call Desequipar(Userindex, .Invent.AnilloEqpSlot)
                    End If
                
588                 .Invent.Object(slot).Equipped = 1
590                 .Invent.AnilloEqpObjIndex = .Invent.Object(slot).ObjIndex
592                 .Invent.AnilloEqpSlot = slot
                
594                 If Len(obj.CreaGRH) <> 0 Then
596                     .Char.Anillo_Aura = obj.CreaGRH
598                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Anillo_Aura, False, 6))
                    End If

600                 If obj.MagicDamageBonus > 0 Then
602                     Call WriteUpdateDM(Userindex)
                    End If
                
604                 If obj.ResistenciaMagica > 0 Then
606                     Call WriteUpdateRM(Userindex)
                    End If

            End Select
    
        End With

        'Actualiza
608     Call UpdateUserInv(False, Userindex, slot)

        Exit Sub
    
ErrHandler:
610     Debug.Print errordesc
612     Call LogError("EquiparInvItem Slot:" & slot & " - Error: " & Err.Number & " - Error Description : " & Err.description & "- " & errordesc)

End Sub

Public Function CheckRazaUsaRopa(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(Userindex) Then
102         CheckRazaUsaRopa = True
            Exit Function

        End If

104     Select Case UserList(Userindex).raza

            Case eRaza.Humano

106             If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 And ObjData(ItemIndex).RazaDrow = 0 Then
108                 If ObjData(ItemIndex).Ropaje > 0 Then
110                     CheckRazaUsaRopa = True
                        Exit Function

                    End If

                End If

112         Case eRaza.Elfo

114             If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 And ObjData(ItemIndex).RazaDrow = 0 Then
116                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
118         Case eRaza.Orco

120             If ObjData(ItemIndex).RazaEnana = 0 Then
122                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
124         Case eRaza.Drow

126             If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 Then
128                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
130         Case eRaza.Gnomo

132             If ObjData(ItemIndex).RazaEnana > 0 Then
134                 CheckRazaUsaRopa = True
                    Exit Function

                End If
        
136         Case eRaza.Enano

138             If ObjData(ItemIndex).RazaEnana > 0 Then
140                 CheckRazaUsaRopa = True
                    Exit Function

                End If
    
        End Select

142     CheckRazaUsaRopa = False

        Exit Function
ErrHandler:
144     Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Public Function CheckRazaTipo(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(Userindex) Then

102         CheckRazaTipo = True
            Exit Function

        End If

104     Select Case ObjData(ItemIndex).RazaTipo

            Case 0
106             CheckRazaTipo = True

108         Case 1

110             If UserList(Userindex).raza = eRaza.Elfo Then
112                 CheckRazaTipo = True
                    Exit Function

                End If
        
114             If UserList(Userindex).raza = eRaza.Drow Then
116                 CheckRazaTipo = True
                    Exit Function

                End If
        
118             If UserList(Userindex).raza = eRaza.Humano Then
120                 CheckRazaTipo = True
                    Exit Function

                End If

122         Case 2

124             If UserList(Userindex).raza = eRaza.Gnomo Then CheckRazaTipo = True
126             If UserList(Userindex).raza = eRaza.Enano Then CheckRazaTipo = True
                Exit Function

128         Case 3

130             If UserList(Userindex).raza = eRaza.Orco Then CheckRazaTipo = True
                Exit Function
    
        End Select

        Exit Function
ErrHandler:
132     Call LogError("Error CheckRazaTipo ItemIndex:" & ItemIndex)

End Function

Public Function CheckClaseTipo(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(Userindex) Then

102         CheckClaseTipo = True
            Exit Function

        End If

104     Select Case ObjData(ItemIndex).ClaseTipo

            Case 0
106             CheckClaseTipo = True
                Exit Function

108         Case 2

110             If UserList(Userindex).clase = eClass.Mage Then CheckClaseTipo = True
112             If UserList(Userindex).clase = eClass.Druid Then CheckClaseTipo = True
                Exit Function

114         Case 1

116             If UserList(Userindex).clase = eClass.Warrior Then CheckClaseTipo = True
118             If UserList(Userindex).clase = eClass.Assasin Then CheckClaseTipo = True
120             If UserList(Userindex).clase = eClass.Bard Then CheckClaseTipo = True
122             If UserList(Userindex).clase = eClass.Cleric Then CheckClaseTipo = True
124             If UserList(Userindex).clase = eClass.Paladin Then CheckClaseTipo = True
126             If UserList(Userindex).clase = eClass.Trabajador Then CheckClaseTipo = True
128             If UserList(Userindex).clase = eClass.Hunter Then CheckClaseTipo = True
                Exit Function

        End Select

        Exit Function
ErrHandler:
130     Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal Userindex As Integer, ByVal slot As Byte)

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

100     If UserList(Userindex).Invent.Object(slot).Amount = 0 Then Exit Sub

102     obj = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex)

104     If obj.OBJType = eOBJType.otWeapon Then
106         If obj.proyectil = 1 Then

                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
108             If Not IntervaloPermiteUsar(Userindex, False) Then Exit Sub
            Else

                'dagas
110             If Not IntervaloPermiteUsar(Userindex) Then Exit Sub

            End If

        Else

112         If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
114         If Not IntervaloPermiteGolpeUsar(Userindex, False) Then Exit Sub

        End If

116     If UserList(Userindex).flags.Meditando Then
118         UserList(Userindex).flags.Meditando = False
120         UserList(Userindex).Char.FX = 0
122         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(UserList(Userindex).Char.CharIndex, 0))
        End If

124     If obj.Newbie = 1 And Not EsNewbie(Userindex) And Not EsGM(Userindex) Then
126         Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

128     If UserList(Userindex).Stats.ELV < obj.MinELV Then
130         Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

132     ObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
134     UserList(Userindex).flags.TargetObjInvIndex = ObjIndex
136     UserList(Userindex).flags.TargetObjInvSlot = slot

138     Select Case obj.OBJType

            Case eOBJType.otUseOnce

140             If UserList(Userindex).flags.Muerto = 1 Then
142                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                'Usa el item
144             UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam + obj.MinHam

146             If UserList(Userindex).Stats.MinHam > UserList(Userindex).Stats.MaxHam Then UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam
148             UserList(Userindex).flags.Hambre = 0
150             Call WriteUpdateHungerAndThirst(Userindex)
                'Sonido
        
152             If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
154                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                Else
156                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                End If
        
                'Quitamos del inv el item
158             Call QuitarUserInvItem(Userindex, slot, 1)
        
160             Call UpdateUserInv(False, Userindex, slot)

162         Case eOBJType.otGuita

164             If UserList(Userindex).flags.Muerto = 1 Then
166                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
168             UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + UserList(Userindex).Invent.Object(slot).Amount
170             UserList(Userindex).Invent.Object(slot).Amount = 0
172             UserList(Userindex).Invent.Object(slot).ObjIndex = 0
174             UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
        
176             Call UpdateUserInv(False, Userindex, slot)
178             Call WriteUpdateGold(Userindex)
        
180         Case eOBJType.otWeapon

182             If UserList(Userindex).flags.Muerto = 1 Then
184                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
186             If Not UserList(Userindex).Stats.MinSta > 0 Then
188                 Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
190             If ObjData(ObjIndex).proyectil = 1 Then
                    'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
192                 Call WriteWorkRequestTarget(Userindex, Proyectiles)
                Else

194                 If UserList(Userindex).flags.TargetObj = Leña Then
196                     If UserList(Userindex).Invent.Object(slot).ObjIndex = DAGA Then
198                         Call TratarDeHacerFogata(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY, Userindex)

                        End If

                    End If

                End If
        
                'REVISAR LADDER
                'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
200             If UserList(Userindex).Invent.Object(slot).Equipped = 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
202         Case eOBJType.otHerramientas

204             If UserList(Userindex).flags.Muerto = 1 Then
206                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
208             If Not UserList(Userindex).Stats.MinSta > 0 Then
210                 Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
212             If UserList(Userindex).Invent.Object(slot).Equipped = 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
214                 Call WriteLocaleMsg(Userindex, "376", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

216             Select Case obj.Subtipo
                
                    Case 1, 2  ' Herramientas del Pescador - Caña y Red
218                     Call WriteWorkRequestTarget(Userindex, eSkill.Pescar)
                
220                 Case 3     ' Herramientas de Alquimia - Tijeras
222                     Call WriteWorkRequestTarget(Userindex, eSkill.Alquimia)
                
224                 Case 4     ' Herramientas de Alquimia - Olla
226                     Call EnivarObjConstruiblesAlquimia(Userindex)
228                     Call WriteShowAlquimiaForm(Userindex)
                
230                 Case 5     ' Herramientas de Carpinteria - Serrucho
232                     Call EnivarObjConstruibles(Userindex)
234                     Call WriteShowCarpenterForm(Userindex)
                
236                 Case 6     ' Herramientas de Tala - Hacha
238                     Call WriteWorkRequestTarget(Userindex, eSkill.Talar)

240                 Case 7     ' Herramientas de Herrero - Martillo
242                     Call WriteConsoleMsg(Userindex, "Debes hacer click derecho sobre el yunque.", FontTypeNames.FONTTYPE_INFOIAO)

244                 Case 8     ' Herramientas de Mineria - Piquete
246                     Call WriteWorkRequestTarget(Userindex, eSkill.Mineria)
                
248                 Case 9     ' Herramientas de Sastreria - Costurero
250                     Call EnivarObjConstruiblesSastre(Userindex)
252                     Call WriteShowSastreForm(Userindex)

                End Select
    
254         Case eOBJType.otPociones

256             If UserList(Userindex).flags.Muerto = 1 Then
258                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
260             UserList(Userindex).flags.TomoPocion = True
262             UserList(Userindex).flags.TipoPocion = obj.TipoPocion
                
                Dim CabezaFinal  As Integer

                Dim CabezaActual As Integer

264             Select Case UserList(Userindex).flags.TipoPocion
        
                    Case 1 'Modif la agilidad
266                     UserList(Userindex).flags.DuracionEfecto = obj.DuracionEfecto
        
                        'Usa el item
268                     UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                
270                     If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                    
272                     If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(Userindex).Stats.UserAtributosBackUP(Agilidad) Then UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(Userindex).Stats.UserAtributosBackUP(Agilidad)
                
274                     Call WriteFYA(Userindex)
                
                        'Quitamos del inv el item
276                     Call QuitarUserInvItem(Userindex, slot, 1)

278                     If obj.Snd1 <> 0 Then
280                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        Else
282                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If
        
284                 Case 2 'Modif la fuerza
286                     UserList(Userindex).flags.DuracionEfecto = obj.DuracionEfecto
        
                        'Usa el item
288                     UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                
290                     If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                
292                     If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(Userindex).Stats.UserAtributosBackUP(Fuerza) Then UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(Userindex).Stats.UserAtributosBackUP(Fuerza)
                
                        'Quitamos del inv el item
294                     Call QuitarUserInvItem(Userindex, slot, 1)

296                     If obj.Snd1 <> 0 Then
298                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        Else
300                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

302                     Call WriteFYA(Userindex)

304                 Case 3 'Pocion roja, restaura HP
                
                        'Usa el item
306                     UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)

308                     If UserList(Userindex).Stats.MinHp > UserList(Userindex).Stats.MaxHp Then UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
                
                        'Quitamos del inv el item
310                     Call QuitarUserInvItem(Userindex, slot, 1)

312                     If obj.Snd1 <> 0 Then
314                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                        Else
316                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If
            
318                 Case 4 'Pocion azul, restaura MANA
            
                        Dim porcentajeRec As Byte
320                     porcentajeRec = obj.Porcentaje
                
                        'Usa el item
322                     UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN + Porcentaje(UserList(Userindex).Stats.MaxMAN, porcentajeRec)

324                     If UserList(Userindex).Stats.MinMAN > UserList(Userindex).Stats.MaxMAN Then UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
                
                        'Quitamos del inv el item
326                     Call QuitarUserInvItem(Userindex, slot, 1)

328                     If obj.Snd1 <> 0 Then
330                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                        Else
332                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If
                
334                 Case 5 ' Pocion violeta

336                     If UserList(Userindex).flags.Envenenado > 0 Then
338                         UserList(Userindex).flags.Envenenado = 0
340                         Call WriteConsoleMsg(Userindex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                            'Quitamos del inv el item
342                         Call QuitarUserInvItem(Userindex, slot, 1)

344                         If obj.Snd1 <> 0 Then
346                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                            Else
348                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                            End If

                        Else
350                         Call WriteConsoleMsg(Userindex, "¡No te encuentras envenenado!", FontTypeNames.FONTTYPE_INFO)

                        End If
                
352                 Case 6  ' Remueve Parálisis

354                     If UserList(Userindex).flags.Paralizado = 1 Or UserList(Userindex).flags.Inmovilizado = 1 Then
356                         If UserList(Userindex).flags.Paralizado = 1 Then
358                             UserList(Userindex).flags.Paralizado = 0
360                             Call WriteParalizeOK(Userindex)

                            End If
                        
362                         If UserList(Userindex).flags.Inmovilizado = 1 Then
364                             UserList(Userindex).Counters.Inmovilizado = 0
366                             UserList(Userindex).flags.Inmovilizado = 0
368                             Call WriteInmovilizaOK(Userindex)

                            End If
                        
                        
                        
370                         Call QuitarUserInvItem(Userindex, slot, 1)

372                         If obj.Snd1 <> 0 Then
374                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                            Else
376                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(255, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                            End If

378                         Call WriteConsoleMsg(Userindex, "Te has removido la paralizis.", FontTypeNames.FONTTYPE_INFOIAO)
                        Else
380                         Call WriteConsoleMsg(Userindex, "No estas paralizado.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
382                 Case 7  ' Pocion Naranja
384                     UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)

386                     If UserList(Userindex).Stats.MinSta > UserList(Userindex).Stats.MaxSta Then UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
                    
                        'Quitamos del inv el item
388                     Call QuitarUserInvItem(Userindex, slot, 1)

390                     If obj.Snd1 <> 0 Then
392                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                        Else
394                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

396                 Case 8  ' Pocion cambio cara

398                     Select Case UserList(Userindex).genero

                            Case eGenero.Hombre

400                             Select Case UserList(Userindex).raza

                                    Case eRaza.Humano
402                                     CabezaFinal = RandomNumber(1, 40)

404                                 Case eRaza.Elfo
406                                     CabezaFinal = RandomNumber(101, 132)

408                                 Case eRaza.Drow
410                                     CabezaFinal = RandomNumber(201, 229)

412                                 Case eRaza.Enano
414                                     CabezaFinal = RandomNumber(301, 329)

416                                 Case eRaza.Gnomo
418                                     CabezaFinal = RandomNumber(401, 429)

420                                 Case eRaza.Orco
422                                     CabezaFinal = RandomNumber(501, 529)

                                End Select

424                         Case eGenero.Mujer

426                             Select Case UserList(Userindex).raza

                                    Case eRaza.Humano
428                                     CabezaFinal = RandomNumber(50, 80)

430                                 Case eRaza.Elfo
432                                     CabezaFinal = RandomNumber(150, 179)

434                                 Case eRaza.Drow
436                                     CabezaFinal = RandomNumber(250, 279)

438                                 Case eRaza.Gnomo
440                                     CabezaFinal = RandomNumber(350, 379)

442                                 Case eRaza.Enano
444                                     CabezaFinal = RandomNumber(450, 479)

446                                 Case eRaza.Orco
448                                     CabezaFinal = RandomNumber(550, 579)

                                End Select

                        End Select
            
450                     UserList(Userindex).Char.Head = CabezaFinal
452                     UserList(Userindex).OrigChar.Head = CabezaFinal
454                     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, CabezaFinal, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                        'Quitamos del inv el item
456                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 102, 0))

458                     If CabezaActual <> CabezaFinal Then
460                         Call QuitarUserInvItem(Userindex, slot, 1)
                        Else
462                         Call WriteConsoleMsg(Userindex, "¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

464                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
466                 Case 9  ' Pocion sexo
    
468                     Select Case UserList(Userindex).genero

                            Case eGenero.Hombre
470                             UserList(Userindex).genero = eGenero.Mujer
                    
472                         Case eGenero.Mujer
474                             UserList(Userindex).genero = eGenero.Hombre
                    
                        End Select
            
476                     Select Case UserList(Userindex).genero

                            Case eGenero.Hombre

478                             Select Case UserList(Userindex).raza

                                    Case eRaza.Humano
480                                     CabezaFinal = RandomNumber(1, 40)

482                                 Case eRaza.Elfo
484                                     CabezaFinal = RandomNumber(101, 132)

486                                 Case eRaza.Drow
488                                     CabezaFinal = RandomNumber(201, 229)

490                                 Case eRaza.Enano
492                                     CabezaFinal = RandomNumber(301, 329)

494                                 Case eRaza.Gnomo
496                                     CabezaFinal = RandomNumber(401, 429)

498                                 Case eRaza.Orco
500                                     CabezaFinal = RandomNumber(501, 529)

                                End Select

502                         Case eGenero.Mujer

504                             Select Case UserList(Userindex).raza

                                    Case eRaza.Humano
506                                     CabezaFinal = RandomNumber(50, 80)

508                                 Case eRaza.Elfo
510                                     CabezaFinal = RandomNumber(150, 179)

512                                 Case eRaza.Drow
514                                     CabezaFinal = RandomNumber(250, 279)

516                                 Case eRaza.Gnomo
518                                     CabezaFinal = RandomNumber(350, 379)

520                                 Case eRaza.Enano
522                                     CabezaFinal = RandomNumber(450, 479)

524                                 Case eRaza.Orco
526                                     CabezaFinal = RandomNumber(550, 579)

                                End Select

                        End Select
            
528                     UserList(Userindex).Char.Head = CabezaFinal
530                     UserList(Userindex).OrigChar.Head = CabezaFinal
532                     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, CabezaFinal, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                        'Quitamos del inv el item
534                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 102, 0))
536                     Call QuitarUserInvItem(Userindex, slot, 1)

538                     If obj.Snd1 <> 0 Then
540                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        Else
542                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If
                
544                 Case 10  ' Invisibilidad
            
546                     If UserList(Userindex).flags.invisible = 0 Then
548                         UserList(Userindex).flags.invisible = 1
550                         UserList(Userindex).Counters.Invisibilidad = obj.DuracionEfecto
552                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, True))
554                         Call WriteContadores(Userindex)
556                         Call QuitarUserInvItem(Userindex, slot, 1)

558                         If obj.Snd1 <> 0 Then
560                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                            Else
562                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("123", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                            End If

564                         Call WriteConsoleMsg(Userindex, "Te has escondido entre las sombras...", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        
                        Else
566                         Call WriteConsoleMsg(Userindex, "Ya estas invisible.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            Exit Sub

                        End If
                    
568                 Case 11  ' Experiencia

                        Dim HR   As Integer

                        Dim MS   As Integer

                        Dim SS   As Integer

                        Dim secs As Integer

570                     If UserList(Userindex).flags.ScrollExp = 1 Then
572                         UserList(Userindex).flags.ScrollExp = obj.CuantoAumento
574                         UserList(Userindex).Counters.ScrollExperiencia = obj.DuracionEfecto
576                         Call QuitarUserInvItem(Userindex, slot, 1)
                        
578                         secs = obj.DuracionEfecto
580                         HR = secs \ 3600
582                         MS = (secs Mod 3600) \ 60
584                         SS = (secs Mod 3600) Mod 60

586                         If SS > 9 Then
588                             Call WriteConsoleMsg(Userindex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                            Else
590                             Call WriteConsoleMsg(Userindex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)

                            End If

                        Else
592                         Call WriteConsoleMsg(Userindex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                            Exit Sub

                        End If

594                     Call WriteContadores(Userindex)

596                     If obj.Snd1 <> 0 Then
598                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                        Else
600                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

602                 Case 12  ' Oro
            
604                     If UserList(Userindex).flags.ScrollOro = 1 Then
606                         UserList(Userindex).flags.ScrollOro = obj.CuantoAumento
608                         UserList(Userindex).Counters.ScrollOro = obj.DuracionEfecto
610                         Call QuitarUserInvItem(Userindex, slot, 1)
612                         secs = obj.DuracionEfecto
614                         HR = secs \ 3600
616                         MS = (secs Mod 3600) \ 60
618                         SS = (secs Mod 3600) Mod 60

620                         If SS > 9 Then
622                             Call WriteConsoleMsg(Userindex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                            Else
624                             Call WriteConsoleMsg(Userindex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)

                            End If
                        
                        Else
626                         Call WriteConsoleMsg(Userindex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                            Exit Sub

                        End If

628                     Call WriteContadores(Userindex)

630                     If obj.Snd1 <> 0 Then
632                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                        Else
634                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

636                 Case 13
                
638                     Call QuitarUserInvItem(Userindex, slot, 1)
640                     UserList(Userindex).flags.Envenenado = 0
642                     UserList(Userindex).flags.Incinerado = 0
                    
644                     If UserList(Userindex).flags.Inmovilizado = 1 Then
646                         UserList(Userindex).Counters.Inmovilizado = 0
648                         UserList(Userindex).flags.Inmovilizado = 0
650                         Call WriteInmovilizaOK(Userindex)
                        

                        End If
                    
652                     If UserList(Userindex).flags.Paralizado = 1 Then
654                         UserList(Userindex).flags.Paralizado = 0
656                         Call WriteParalizeOK(Userindex)
                        

                        End If
                    
658                     If UserList(Userindex).flags.Ceguera = 1 Then
660                         UserList(Userindex).flags.Ceguera = 0
662                         Call WriteBlindNoMore(Userindex)
                        

                        End If
                    
664                     If UserList(Userindex).flags.Maldicion = 1 Then
666                         UserList(Userindex).flags.Maldicion = 0
668                         UserList(Userindex).Counters.Maldicion = 0

                        End If
                    
670                     UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
672                     UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
674                     UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
676                     UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
678                     UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam
                    
680                     UserList(Userindex).flags.Hambre = 0
682                     UserList(Userindex).flags.Sed = 0
                    
684                     Call WriteUpdateHungerAndThirst(Userindex)
686                     Call WriteConsoleMsg(Userindex, "Donador> Te sentis sano y lleno.", FontTypeNames.FONTTYPE_WARNING)

688                     If obj.Snd1 <> 0 Then
690                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                        Else
692                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

694                 Case 14
                
696                     If UserList(Userindex).flags.BattleModo = 1 Then
698                         Call WriteConsoleMsg(Userindex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub

                        End If
                    
700                     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = CARCEL Then
702                         Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                    
                        Dim Map     As Integer

                        Dim X       As Byte

                        Dim Y       As Byte

                        Dim DeDonde As WorldPos

704                     Call QuitarUserInvItem(Userindex, slot, 1)
            
706                     Select Case UserList(Userindex).Hogar

                            Case eCiudad.cUllathorpe
708                             DeDonde = Ullathorpe
                            
710                         Case eCiudad.cNix
712                             DeDonde = Nix
                
714                         Case eCiudad.cBanderbill
716                             DeDonde = Banderbill
                        
718                         Case eCiudad.cLindos
720                             DeDonde = Lindos
                            
722                         Case eCiudad.cArghal
724                             DeDonde = Arghal
                            
726                         Case eCiudad.CHillidan
728                             DeDonde = Hillidan
                            
730                         Case Else
732                             DeDonde = Ullathorpe

                        End Select
                    
734                     Map = DeDonde.Map
736                     X = DeDonde.X
738                     Y = DeDonde.Y
                    
740                     Call FindLegalPos(Userindex, Map, X, Y)
742                     Call WarpUserChar(Userindex, Map, X, Y, True)
744                     Call WriteConsoleMsg(Userindex, "Ya estas a salvo...", FontTypeNames.FONTTYPE_WARNING)

746                     If obj.Snd1 <> 0 Then
748                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                        Else
750                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

752                 Case 15  ' Aliento de sirena
                        
754                     If UserList(Userindex).Counters.Oxigeno >= 3540 Then
                        
756                         Call WriteConsoleMsg(Userindex, "No podes acumular más de 59 minutos de oxigeno.", FontTypeNames.FONTTYPE_INFOIAO)
758                         secs = UserList(Userindex).Counters.Oxigeno
760                         HR = secs \ 3600
762                         MS = (secs Mod 3600) \ 60
764                         SS = (secs Mod 3600) Mod 60

766                         If SS > 9 Then
768                             Call WriteConsoleMsg(Userindex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                            Else
770                             Call WriteConsoleMsg(Userindex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)

                            End If

                        Else
                            
772                         UserList(Userindex).Counters.Oxigeno = UserList(Userindex).Counters.Oxigeno + obj.DuracionEfecto
774                         Call QuitarUserInvItem(Userindex, slot, 1)
                            
                            'secs = UserList(UserIndex).Counters.Oxigeno
                            ' HR = secs \ 3600
                            ' MS = (secs Mod 3600) \ 60
                            ' SS = (secs Mod 3600) Mod 60
                            ' If SS > 9 Then
                            ' Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                            'Call WriteConsoleMsg(UserIndex, "Has agregado " & MS & ":" & SS & " segundos de oxigeno.", FontTypeNames.FONTTYPE_New_DONADOR)
                            ' Else
                            ' Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)
                            ' End If
                            
776                         UserList(Userindex).flags.Ahogandose = 0
778                         Call WriteOxigeno(Userindex)
                            
780                         Call WriteContadores(Userindex)

782                         If obj.Snd1 <> 0 Then
784                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                            Else
786                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                            End If

                        End If

788                 Case 16 ' Divorcio

790                     If UserList(Userindex).flags.Casado = 1 Then

                            Dim tUser As Integer

                            'UserList(UserIndex).flags.Pareja
792                         tUser = NameIndex(UserList(Userindex).flags.Pareja)
794                         Call QuitarUserInvItem(Userindex, slot, 1)
                        
796                         If tUser <= 0 Then

                                Dim FileUser As String

798                             FileUser = CharPath & UCase$(UserList(Userindex).flags.Pareja) & ".chr"
                                'Call WriteVar(FileUser, "FLAGS", "CASADO", 0)
                                'Call WriteVar(FileUser, "FLAGS", "PAREJA", "")
800                             UserList(Userindex).flags.Casado = 0
802                             UserList(Userindex).flags.Pareja = ""
804                             Call WriteConsoleMsg(Userindex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
806                             UserList(Userindex).MENSAJEINFORMACION = UserList(Userindex).name & " se ha divorciado de ti."

                            Else
808                             UserList(tUser).flags.Casado = 0
810                             UserList(tUser).flags.Pareja = ""
812                             UserList(Userindex).flags.Casado = 0
814                             UserList(Userindex).flags.Pareja = ""
816                             Call WriteConsoleMsg(Userindex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
818                             Call WriteConsoleMsg(tUser, UserList(Userindex).name & " se ha divorciado de ti.", FontTypeNames.FONTTYPE_INFOIAO)
                            
                            End If

820                         If obj.Snd1 <> 0 Then
822                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                            Else
824                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                            End If
                    
                        Else
826                         Call WriteConsoleMsg(Userindex, "No estas casado.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

828                 Case 17 'Cara legendaria

830                     Select Case UserList(Userindex).genero

                            Case eGenero.Hombre

832                             Select Case UserList(Userindex).raza

                                    Case eRaza.Humano
834                                     CabezaFinal = RandomNumber(684, 686)

836                                 Case eRaza.Elfo
838                                     CabezaFinal = RandomNumber(690, 692)

840                                 Case eRaza.Drow
842                                     CabezaFinal = RandomNumber(696, 698)

844                                 Case eRaza.Enano
846                                     CabezaFinal = RandomNumber(702, 704)

848                                 Case eRaza.Gnomo
850                                     CabezaFinal = RandomNumber(708, 710)

852                                 Case eRaza.Orco
854                                     CabezaFinal = RandomNumber(714, 716)

                                End Select

856                         Case eGenero.Mujer

858                             Select Case UserList(Userindex).raza

                                    Case eRaza.Humano
860                                     CabezaFinal = RandomNumber(687, 689)

862                                 Case eRaza.Elfo
864                                     CabezaFinal = RandomNumber(693, 695)

866                                 Case eRaza.Drow
868                                     CabezaFinal = RandomNumber(699, 701)

870                                 Case eRaza.Gnomo
872                                     CabezaFinal = RandomNumber(705, 707)

874                                 Case eRaza.Enano
876                                     CabezaFinal = RandomNumber(711, 713)

878                                 Case eRaza.Orco
880                                     CabezaFinal = RandomNumber(717, 719)

                                End Select

                        End Select

882                     CabezaActual = UserList(Userindex).OrigChar.Head
                        
884                     UserList(Userindex).Char.Head = CabezaFinal
886                     UserList(Userindex).OrigChar.Head = CabezaFinal
888                     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, CabezaFinal, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)

                        'Quitamos del inv el item
890                     If CabezaActual <> CabezaFinal Then
892                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 102, 0))
894                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
896                         Call QuitarUserInvItem(Userindex, slot, 1)
                        Else
898                         Call WriteConsoleMsg(Userindex, "¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

900                 Case 18  ' tan solo crea una particula por determinado tiempo

                        Dim Particula           As Integer

                        Dim Tiempo              As Long

                        Dim ParticulaPermanente As Byte

                        Dim sobrechar           As Byte

902                     If obj.CreaParticula <> "" Then
904                         Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
906                         Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
908                         ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
910                         sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                            
912                         If ParticulaPermanente = 1 Then
914                             UserList(Userindex).Char.ParticulaFx = Particula
916                             UserList(Userindex).Char.loops = Tiempo

                            End If
                            
918                         If sobrechar = 1 Then
920                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, Particula, Tiempo))
                            Else
                            
922                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, Particula, Tiempo, False))

                                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, Particula, Tiempo))
                            End If

                        End If
                        
924                     If obj.CreaFX <> 0 Then
926                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageFxPiso(obj.CreaFX, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, obj.CreaFX, 0))
                            ' PrepareMessageCreateFX
                        End If
                        
928                     If obj.Snd1 <> 0 Then
930                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If
                        
932                     Call QuitarUserInvItem(Userindex, slot, 1)

934                 Case 19 ' Reseteo de skill

                        Dim S As Byte
                
936                     If UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) >= 80 Then
938                         Call WriteConsoleMsg(Userindex, "Has fundado un clan, no podes resetar tus skills. ", FontTypeNames.FONTTYPE_INFOIAO)
                            Exit Sub

                        End If
                    
940                     For S = 1 To NUMSKILLS
942                         UserList(Userindex).Stats.UserSkills(S) = 0
944                     Next S
                    
                        Dim SkillLibres As Integer
                    
946                     SkillLibres = 5
948                     SkillLibres = SkillLibres + (5 * UserList(Userindex).Stats.ELV)
                     
950                     UserList(Userindex).Stats.SkillPts = SkillLibres
952                     Call WriteLevelUp(Userindex, UserList(Userindex).Stats.SkillPts)
                    
954                     Call WriteConsoleMsg(Userindex, "Tus skills han sido reseteados.", FontTypeNames.FONTTYPE_INFOIAO)
956                     Call QuitarUserInvItem(Userindex, slot, 1)

958                 Case 20
                
960                     If UserList(Userindex).Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
962                         UserList(Userindex).Stats.InventLevel = UserList(Userindex).Stats.InventLevel + 1
964                         UserList(Userindex).CurrentInventorySlots = getMaxInventorySlots(Userindex)
966                         Call WriteInventoryUnlockSlots(Userindex)
968                         Call WriteConsoleMsg(Userindex, "Has aumentado el espacio de tu inventario!", FontTypeNames.FONTTYPE_INFO)
970                         Call QuitarUserInvItem(Userindex, slot, 1)
                        Else
972                         Call WriteConsoleMsg(Userindex, "Ya has desbloqueado todos los casilleros disponibles.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                
                End Select

974             Call WriteUpdateUserStats(Userindex)
976             Call UpdateUserInv(False, Userindex, slot)

978         Case eOBJType.otBebidas

980             If UserList(Userindex).flags.Muerto = 1 Then
982                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

984             UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + obj.MinSed

986             If UserList(Userindex).Stats.MinAGU > UserList(Userindex).Stats.MaxAGU Then UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
988             UserList(Userindex).flags.Sed = 0
990             Call WriteUpdateHungerAndThirst(Userindex)
        
                'Quitamos del inv el item
992             Call QuitarUserInvItem(Userindex, slot, 1)
        
994             If obj.Snd1 <> 0 Then
996                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
            
                Else
998                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                 End If
        
1000             Call UpdateUserInv(False, Userindex, slot)
        
1002         Case eOBJType.OtCofre

1004             If UserList(Userindex).flags.Muerto = 1 Then
1006                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

                 'Quitamos del inv el item
1008             Call QuitarUserInvItem(Userindex, slot, 1)
1010             Call UpdateUserInv(False, Userindex, slot)
        
1012             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg(UserList(Userindex).name & " ha abierto un " & obj.name & " y obtuvo...", FontTypeNames.FONTTYPE_New_DONADOR))
        
1014             If obj.Snd1 <> 0 Then
1016                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                 End If
        
1018             If obj.CreaFX <> 0 Then
1020                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, obj.CreaFX, 0))

                 End If
        
                 Dim i As Byte

1022             If obj.Subtipo = 1 Then

1024                 For i = 1 To obj.CantItem

1026                     If Not MeterItemEnInventario(Userindex, obj.Item(i)) Then Call TirarItemAlPiso(UserList(Userindex).Pos, obj.Item(i))
1028                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
1030                 Next i
        
                 Else
        
1032                 For i = 1 To obj.CantEntrega

                         Dim indexobj As Byte
                
1034                     indexobj = RandomNumber(1, obj.CantItem)
            
                         Dim Index As obj

1036                     Index.ObjIndex = obj.Item(indexobj).ObjIndex
1038                     Index.Amount = obj.Item(indexobj).Amount

1040                     If Not MeterItemEnInventario(Userindex, Index) Then Call TirarItemAlPiso(UserList(Userindex).Pos, Index)
1042                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).name & " (" & Index.Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
1044                 Next i

                 End If
    
1046         Case eOBJType.otLlaves
1048             Call WriteConsoleMsg(Userindex, "Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.", FontTypeNames.FONTTYPE_INFO)
    
1050         Case eOBJType.otBotellaVacia

1052             If UserList(Userindex).flags.Muerto = 1 Then
1054                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1056             If (MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
1058                 Call WriteConsoleMsg(Userindex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1060             MiObj.Amount = 1
1062             MiObj.ObjIndex = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex).IndexAbierta
1064             Call QuitarUserInvItem(Userindex, slot, 1)

1066             If Not MeterItemEnInventario(Userindex, MiObj) Then
1068                 Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

                 End If
        
1070             Call UpdateUserInv(False, Userindex, slot)
    
1072         Case eOBJType.otBotellaLlena

1074             If UserList(Userindex).flags.Muerto = 1 Then
1076                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1078             UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + obj.MinSed

1080             If UserList(Userindex).Stats.MinAGU > UserList(Userindex).Stats.MaxAGU Then UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
1082             UserList(Userindex).flags.Sed = 0
1084             Call WriteUpdateHungerAndThirst(Userindex)
1086             MiObj.Amount = 1
1088             MiObj.ObjIndex = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex).IndexCerrada
1090             Call QuitarUserInvItem(Userindex, slot, 1)

1092             If Not MeterItemEnInventario(Userindex, MiObj) Then
1094                 Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

                 End If
        
1096             Call UpdateUserInv(False, Userindex, slot)
    
1098         Case eOBJType.otPergaminos

1100             If UserList(Userindex).flags.Muerto = 1 Then
1102                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
                 'Call LogError(UserList(UserIndex).Name & " intento aprender el hechizo " & ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex)
        
1104             If ClasePuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(slot).ObjIndex, slot) Then

                     'If UserList(UserIndex).Stats.MaxMAN > 0 Then
1106                 If UserList(Userindex).flags.Hambre = 0 And UserList(Userindex).flags.Sed = 0 Then
1108                     Call AgregarHechizo(Userindex, slot)
1110                     Call UpdateUserInv(False, Userindex, slot)
                         ' Call LogError(UserList(UserIndex).Name & " lo aprendio.")
                     Else
1112                     Call WriteConsoleMsg(Userindex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                     End If

                     ' Else
                     '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_WARNING)
                     'End If
                 Else
             
1114                 Call WriteConsoleMsg(Userindex, "Por mas que lo intentas, no podés comprender el manuescrito.", FontTypeNames.FONTTYPE_INFO)
   
                 End If
        
1116         Case eOBJType.otMinerales

1118             If UserList(Userindex).flags.Muerto = 1 Then
1120                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1122             Call WriteWorkRequestTarget(Userindex, FundirMetal)
       
1124         Case eOBJType.otInstrumentos

1126             If UserList(Userindex).flags.Muerto = 1 Then
1128                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1130             If obj.Real Then '¿Es el Cuerno Real?
1132                 If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
1134                     If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
1136                         Call WriteConsoleMsg(Userindex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If

1138                     Call SendData(SendTarget.toMap, UserList(Userindex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                         Exit Sub
                     Else
1140                     Call WriteConsoleMsg(Userindex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                         Exit Sub

                     End If

1142             ElseIf obj.Caos Then '¿Es el Cuerno Legión?

1144                 If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
1146                     If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
1148                         Call WriteConsoleMsg(Userindex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If

1150                     Call SendData(SendTarget.toMap, UserList(Userindex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                         Exit Sub
                     Else
1152                     Call WriteConsoleMsg(Userindex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                         Exit Sub

                     End If

                 End If

                 'Si llega aca es porque es o Laud o Tambor o Flauta
1154             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
       
1156         Case eOBJType.otBarcos
            
                 'Verifica si tiene el nivel requerido para navegar, siendo Trabajador o Pirata
1158             If UserList(Userindex).Stats.ELV < 20 And (UserList(Userindex).clase = eClass.Trabajador Or UserList(Userindex).clase = eClass.Pirat) Then
1160                 Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 'Verifica si tiene el nivel requerido para navegar, sin ser Trabajador o Pirata
1162             ElseIf UserList(Userindex).Stats.ELV < 25 Then
1164                 Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
                 'If obj.Subtipo = 0 Then
1166             If UserList(Userindex).flags.Navegando = 0 Then
1168                 If ((LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X - 1, UserList(Userindex).Pos.Y, True, False) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1, True, False) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X + 1, UserList(Userindex).Pos.Y, True, False) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, True, False)) And UserList(Userindex).flags.Navegando = 0) Or UserList(Userindex).flags.Navegando = 1 Then
1170                     Call DoNavega(Userindex, obj, slot)
                     Else
1172                     Call WriteConsoleMsg(Userindex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)

                     End If

                 Else 'Ladder 10-02-2010

1174                 If UserList(Userindex).Invent.BarcoObjIndex <> UserList(Userindex).Invent.Object(slot).ObjIndex Then
1176                     Call DoReNavega(Userindex, obj, slot)
                     Else

1178                     If ((LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X - 1, UserList(Userindex).Pos.Y, False, True) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1, False, True) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X + 1, UserList(Userindex).Pos.Y, False, True) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, False, True)) And UserList(Userindex).flags.Navegando = 1) Or UserList(Userindex).flags.Navegando = 0 Then
1180                         Call DoNavega(Userindex, obj, slot)
                         Else
1182                         Call WriteConsoleMsg(Userindex, "¡Debes aproximarte a la costa para dejar la barca!", FontTypeNames.FONTTYPE_INFO)

                         End If

                     End If

                 End If

                 'Else
    
                 ' End If
        
                 ' Case eOBJType.otTrajeDeBaño
        
                 '  Dim Puede As Boolean
                 ' Debug.Print "poner traje"
                 ' If UserList(UserIndex).flags.Nadando = 0 Then
        
                 '  If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y).trigger = 8 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1).trigger = 8 Then
                 '      Puede = True
                 ' End If
    
                 '  If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y).trigger = 8 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1).trigger = 8 Then
                 '     Puede = True
                 ' End If
        
                 '   If Puede Then
                 '  Call WriteConsoleMsg(UserIndex, "¡Te hago nadar!", FontTypeNames.FONTTYPE_INFO)
                 '  Call DoNadar(UserIndex, ObjData(UserList(UserIndex).Invent.BarcoObjIndex), 0)
                 ' Call DoNavega(UserIndex, ObjData(UserList(UserIndex).Invent.BarcoObjIndex), UserList(UserIndex).Invent.BarcoSlot)
            
                 'UserList(UserIndex).flags.Nadando = 1
            
                 ' Else
                 '   Call WriteConsoleMsg(UserIndex, "¡No podes nadar!", FontTypeNames.FONTTYPE_INFO)
                 ' End If
        
1184         Case eOBJType.otMonturas
                 'Verifica todo lo que requiere la montura
    
1186             If UserList(Userindex).flags.Muerto = 1 Then
1188                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡Estas muerto! Los fantasmas no pueden montar.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
            
1190             If UserList(Userindex).flags.Navegando = 1 Then
1192                 Call WriteConsoleMsg(Userindex, "Debes dejar de navegar para poder montarté.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1194             If MapInfo(UserList(Userindex).Pos.Map).zone = "DUNGEON" Then
1196                 Call WriteConsoleMsg(Userindex, "No podes cabalgar dentro de un dungeon.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1198             Call DoMontar(Userindex, obj, slot)

1200         Case eOBJType.OtDonador

1202             Select Case obj.Subtipo

                     Case 1
            
1204                     If UserList(Userindex).Counters.Pena <> 0 Then
1206                         Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If
                
1208                     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = CARCEL Then
1210                         Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If
            
1212                     Call WarpUserChar(Userindex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1214                     Call WriteConsoleMsg(Userindex, "Has viajado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
1216                     Call QuitarUserInvItem(Userindex, slot, 1)
1218                     Call UpdateUserInv(False, Userindex, slot)
                
1220                 Case 2

1222                     If DonadorCheck(UserList(Userindex).Cuenta) = 0 Then
1224                         Call DonadorTiempo(UserList(Userindex).Cuenta, CLng(obj.CuantoAumento))
1226                         Call WriteConsoleMsg(Userindex, "Donación> Se han agregado " & obj.CuantoAumento & " dias de donador a tu cuenta. Relogea tu personaje para empezar a disfrutar la experiencia.", FontTypeNames.FONTTYPE_WARNING)
1228                         Call QuitarUserInvItem(Userindex, slot, 1)
1230                         Call UpdateUserInv(False, Userindex, slot)
                         Else
1232                         Call DonadorTiempo(UserList(Userindex).Cuenta, CLng(obj.CuantoAumento))
1234                         Call WriteConsoleMsg(Userindex, "¡Se han añadido " & CLng(obj.CuantoAumento) & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
1236                         UserList(Userindex).donador.activo = 1
1238                         Call QuitarUserInvItem(Userindex, slot, 1)
1240                         Call UpdateUserInv(False, Userindex, slot)

                             'Call WriteConsoleMsg(UserIndex, "Donación> Debes esperar a que finalice el periodo existente para renovar tu suscripción.", FontTypeNames.FONTTYPE_INFOIAO)
                         End If

1242                 Case 3
1244                     Call AgregarCreditosDonador(UserList(Userindex).Cuenta, CLng(obj.CuantoAumento))
1246                     Call WriteConsoleMsg(Userindex, "Donación> Tu credito ahora es de " & CreditosDonadorCheck(UserList(Userindex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
1248                     Call QuitarUserInvItem(Userindex, slot, 1)
1250                     Call UpdateUserInv(False, Userindex, slot)

                 End Select
     
1252         Case eOBJType.otpasajes

1254             If UserList(Userindex).flags.Muerto = 1 Then
1256                 Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1258             If UserList(Userindex).flags.TargetNpcTipo <> Pirata Then
1260                 Call WriteConsoleMsg(Userindex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1262             If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 3 Then
1264                 Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1266             If UserList(Userindex).Pos.Map <> obj.DesdeMap Then
                     Rem  Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aquí! Largate!", FontTypeNames.FONTTYPE_INFO)
1268                 Call WriteChatOverHead(Userindex, "El pasaje no lo compraste aquí! Largate!", str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
                     Exit Sub

                 End If
        
1270             If Not MapaValido(obj.HastaMap) Then
                     Rem Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
1272                 Call WriteChatOverHead(Userindex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
                     Exit Sub

                 End If

1274             If obj.NecesitaNave > 0 Then
1276                 If UserList(Userindex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                         Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
1278                     Call WriteChatOverHead(Userindex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
                         Exit Sub

                     End If

                 End If
            
1280             Call WarpUserChar(Userindex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1282             Call WriteConsoleMsg(Userindex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
1284             UserList(Userindex).Stats.MinAGU = 0
1286             UserList(Userindex).Stats.MinHam = 0
1288             UserList(Userindex).flags.Sed = 1
1290             UserList(Userindex).flags.Hambre = 1
1292             Call WriteUpdateHungerAndThirst(Userindex)
1294             Call QuitarUserInvItem(Userindex, slot, 1)
1296             Call UpdateUserInv(False, Userindex, slot)
        
1298         Case eOBJType.otRunas
    
1300             If UserList(Userindex).Counters.Pena <> 0 Then
1302                 Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1304             If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = CARCEL Then
1306                 Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1308             If UserList(Userindex).flags.BattleModo = 1 Then
1310                 Call WriteConsoleMsg(Userindex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                     Exit Sub

                 End If
        
1312             If MapInfo(UserList(Userindex).Pos.Map).Seguro = 0 And UserList(Userindex).flags.Muerto = 0 Then
1314                 Call WriteConsoleMsg(Userindex, "Solo podes usar tu runa en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1316             If UserList(Userindex).Accion.AccionPendiente Then
                     Exit Sub

                 End If
        
1318             Select Case ObjData(ObjIndex).TipoRuna
        
                     Case 1, 2

1320                     If UserList(Userindex).donador.activo = 0 Then ' Donador no espera tiempo
1322                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
1324                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 350, Accion_Barra.Runa))
                         Else
1326                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
1328                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 100, Accion_Barra.Runa))

                         End If

1330                     UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
1332                     UserList(Userindex).Accion.AccionPendiente = True
1334                     UserList(Userindex).Accion.TipoAccion = Accion_Barra.Runa
1336                     UserList(Userindex).Accion.RunaObj = ObjIndex
1338                     UserList(Userindex).Accion.ObjSlot = slot
            
1340                 Case 3
        
                         Dim parejaindex As Integer

1342                     If Not UserList(Userindex).flags.BattleModo Then
                
                             'If UserList(UserIndex).donador.activo = 1 Then
1344                         If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
1346                             If UserList(Userindex).flags.Casado = 1 Then
1348                                 parejaindex = NameIndex(UserList(Userindex).flags.Pareja)
                        
1350                                 If parejaindex > 0 Then
1352                                     If UserList(parejaindex).flags.BattleModo = 0 Then
1354                                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 600, False))
1356                                         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 600, Accion_Barra.GoToPareja))
1358                                         UserList(Userindex).Accion.AccionPendiente = True
1360                                         UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
1362                                         UserList(Userindex).Accion.TipoAccion = Accion_Barra.GoToPareja
                                         Else
1364                                         Call WriteConsoleMsg(Userindex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                         End If
                                
                                     Else
1366                                     Call WriteConsoleMsg(Userindex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                                     End If

                                 Else
1368                                 Call WriteConsoleMsg(Userindex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                                 End If

                             Else
1370                             Call WriteConsoleMsg(Userindex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                             End If
                
                             ' Else
                             '  Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                             ' End If
                         Else
1372                         Call WriteConsoleMsg(Userindex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
        
                         End If
    
                 End Select
        
1374         Case eOBJType.otmapa
1376             Call WriteShowFrmMapa(Userindex)
        
         End Select

         Exit Sub

hErr:
1378     LogError "Error en useinvitem Usuario: " & UserList(Userindex).name & " item:" & obj.name & " index: " & UserList(Userindex).Invent.Object(slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal Userindex As Integer)
        
        On Error GoTo EnivarArmasConstruibles_Err
        

100     Call WriteBlacksmithWeapons(Userindex)

        
        Exit Sub

EnivarArmasConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarArmasConstruibles", Erl)
104     Resume Next
        
End Sub
 
Sub EnivarObjConstruibles(ByVal Userindex As Integer)
        
        On Error GoTo EnivarObjConstruibles_Err
        

100     Call WriteCarpenterObjects(Userindex)

        
        Exit Sub

EnivarObjConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruibles", Erl)
104     Resume Next
        
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal Userindex As Integer)
        
        On Error GoTo EnivarObjConstruiblesAlquimia_Err
        

100     Call WriteAlquimistaObjects(Userindex)

        
        Exit Sub

EnivarObjConstruiblesAlquimia_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)
104     Resume Next
        
End Sub

Sub EnivarObjConstruiblesSastre(ByVal Userindex As Integer)
        
        On Error GoTo EnivarObjConstruiblesSastre_Err
        

100     Call WriteSastreObjects(Userindex)

        
        Exit Sub

EnivarObjConstruiblesSastre_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)
104     Resume Next
        
End Sub

Sub EnivarArmadurasConstruibles(ByVal Userindex As Integer)
        
        On Error GoTo EnivarArmadurasConstruibles_Err
        

100     Call WriteBlacksmithArmors(Userindex)

        
        Exit Sub

EnivarArmadurasConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarArmadurasConstruibles", Erl)
104     Resume Next
        
End Sub

Sub TirarTodo(ByVal Userindex As Integer)

        On Error Resume Next

100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub
102     If UserList(Userindex).flags.BattleModo = 1 Then Exit Sub

104     Call TirarTodosLosItems(Userindex)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemSeCae_Err
        

100     ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> eOBJType.otLlaves And ObjData(Index).OBJType <> eOBJType.otBarcos And ObjData(Index).OBJType <> eOBJType.otMonturas And ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And ObjData(Index).donador = 0 And Not ObjData(Index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.ItemSeCae", Erl)
104     Resume Next
        
End Function

Public Function PirataCaeItem(ByVal Userindex As Integer, ByVal slot As Byte)

100     With UserList(Userindex)
    
102         If .clase = eClass.Pirat Then
            
                ' El pirata con galera no pierde los últimos 6 * (cada 10 niveles; max 1) slots
104             If ObjData(.Invent.BarcoObjIndex).Ropaje = iGalera Then
            
106                 If slot > .CurrentInventorySlots - 6 * min(.Stats.ELV \ 10, 1) Then
                        Exit Function
                    End If
            
                ' Con galeón no pierde los últimos 6 * (cada 10 niveles; max 3) slots
108             ElseIf ObjData(.Invent.BarcoObjIndex).Ropaje = iGaleon Then
            
110                 If slot > .CurrentInventorySlots - 6 * min(.Stats.ELV \ 10, 3) Then
                        Exit Function
                    End If
            
                End If
            
            End If
        
        End With
    
112     PirataCaeItem = True

End Function

Sub TirarTodosLosItems(ByVal Userindex As Integer)
        
        On Error GoTo TirarTodosLosItems_Err
    
        '***************************************************
        'Author: Unknown
        'Last Modification: 12/01/2010 (ZaMa)
        '12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
        '***************************************************
    
        Dim i         As Byte
        Dim NuevaPos  As WorldPos
        Dim MiObj     As obj
        Dim ItemIndex As Integer
    
100     With UserList(Userindex)
    
102         For i = 1 To .CurrentInventorySlots
    
104             ItemIndex = .Invent.Object(i).ObjIndex

106             If ItemIndex > 0 Then

108                 If ItemSeCae(ItemIndex) And PirataCaeItem(Userindex, i) Then
110                     NuevaPos.X = 0
112                     NuevaPos.Y = 0
                
114                     If .flags.CarroMineria = 1 Then
                
116                         If ItemIndex = ORO_MINA Or ItemIndex = PLATA_MINA Or ItemIndex = HIERRO_MINA Then
                       
118                             MiObj.Amount = .Invent.Object(i).Amount * 0.3
120                             MiObj.ObjIndex = ItemIndex
                        
122                             Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                    
124                             If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
126                                 Call DropObj(Userindex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                                End If

                            End If
                    
                        Else
                    
128                         MiObj.Amount = .Invent.Object(i).Amount
130                         MiObj.ObjIndex = ItemIndex
                        
132                         Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                
134                         If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
136                             Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                            End If
                    
                        End If
                
                    End If

                End If
    
138         Next i
    
        End With
 
        Exit Sub

TirarTodosLosItems_Err:
140     Call RegistrarError(Err.Number, Err.description, "InvUsuario.TirarTodosLosItems", Erl)

142     Resume Next
        
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo ItemNewbie_Err
        

100     ItemNewbie = ObjData(ItemIndex).Newbie = 1

        
        Exit Function

ItemNewbie_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.ItemNewbie", Erl)
104     Resume Next
        
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal Userindex As Integer)
        
        On Error GoTo TirarTodosLosItemsNoNewbies_Err
        

        Dim i         As Byte

        Dim NuevaPos  As WorldPos

        Dim MiObj     As obj

        Dim ItemIndex As Integer

100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub

102     For i = 1 To UserList(Userindex).CurrentInventorySlots
104         ItemIndex = UserList(Userindex).Invent.Object(i).ObjIndex

106         If ItemIndex > 0 Then
108             If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
110                 NuevaPos.X = 0
112                 NuevaPos.Y = 0
            
                    'Creo MiObj
114                 MiObj.Amount = UserList(Userindex).Invent.Object(i).ObjIndex
116                 MiObj.ObjIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
118                 Tilelibre UserList(Userindex).Pos, NuevaPos, MiObj, True, True

120                 If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
122                     If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

124     Next i

        
        Exit Sub

TirarTodosLosItemsNoNewbies_Err:
126     Call RegistrarError(Err.Number, Err.description, "InvUsuario.TirarTodosLosItemsNoNewbies", Erl)
128     Resume Next
        
End Sub
