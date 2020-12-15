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

        '17/09/02
        'Agregue que la función se asegure que el objeto no es un barco

        On Error Resume Next

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
112     If UCase$(MapInfo(UserList(UserIndex).Pos.Map).restrict_mode) = "NEWBIE" Then
        
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
144     Call RegistrarError(Err.Number, Err.description, "InvUsuario.QuitarNewbieObj", Erl)
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

130     UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
132     UserList(UserIndex).Invent.AnilloEqpSlot = 0

134     UserList(UserIndex).Invent.NudilloObjIndex = 0
136     UserList(UserIndex).Invent.NudilloSlot = 0

138     UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
140     UserList(UserIndex).Invent.MunicionEqpSlot = 0

142     UserList(UserIndex).Invent.BarcoObjIndex = 0
144     UserList(UserIndex).Invent.BarcoSlot = 0

146     UserList(UserIndex).Invent.MonturaObjIndex = 0
148     UserList(UserIndex).Invent.MonturaSlot = 0

150     UserList(UserIndex).Invent.MagicoObjIndex = 0
152     UserList(UserIndex).Invent.MagicoSlot = 0

        
        Exit Sub

LimpiarInventario_Err:
154     Call RegistrarError(Err.Number, Err.description, "InvUsuario.LimpiarInventario", Erl)
156     Resume Next
        
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
        '***************************************************
        On Error GoTo ErrHandler

        'If Cantidad > 100000 Then Exit Sub
100     If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub

        'SI EL Pjta TIENE ORO LO TIRAMOS
102     If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then

            Dim i     As Byte

            Dim MiObj As obj

            Dim Logs  As Long

            'info debug
            Dim loops As Integer
        
104         Logs = Cantidad

            Dim Extra    As Long

            Dim TeniaOro As Long

106         TeniaOro = UserList(UserIndex).Stats.GLD

108         If Cantidad > 500000 Then 'Para evitar explotar demasiado
110             Extra = Cantidad - 500000
112             Cantidad = 500000

            End If
        
114         Do While (Cantidad > 0)
            
116             If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
118                 MiObj.Amount = MAX_INVENTORY_OBJS
120                 Cantidad = Cantidad - MiObj.Amount
                Else
122                 MiObj.Amount = Cantidad
124                 Cantidad = Cantidad - MiObj.Amount

                End If

126             MiObj.ObjIndex = iORO

                Dim AuxPos As WorldPos
128             If UserList(UserIndex).clase = eClass.Pirat Then
130                 AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                Else
132                 AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                End If
            
134             If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
136                 UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount

                End If
            
                'info debug
138             loops = loops + 1

140             If loops > 100 Then
142                 LogError ("Error en tiraroro")
                    Exit Sub

                End If
            
            Loop
        
144         If EsGM(UserIndex) Then
146             If MiObj.ObjIndex = iORO Then
148                 Call LogGM(UserList(UserIndex).name, "Tiro: " & Logs & " monedas de oro.")
                Else
150                 Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & Logs & " Objeto:" & ObjData(MiObj.ObjIndex).name)

                End If

            End If
        
152         If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
154         If Extra > 0 Then
156             UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra

            End If
    
        End If

        Exit Sub

ErrHandler:

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
118     Call RegistrarError(Err.Number, Err.description, "InvUsuario.QuitarUserInvItem", Erl)
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
118     Call RegistrarError(Err.Number, Err.description, "InvUsuario.UpdateUserInv", Erl)
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
120     Call RegistrarError(Err.Number, Err.description, "InvUsuario.MakeObj", Erl)

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

140     WriteUpdateGold (UserIndex)
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
146     Call RegistrarError(Err.Number, Err.description, "InvUsuario.GetObj", Erl)
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
       
154         Case eOBJType.otmagicos
    
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
                
322         Case eOBJType.otAnillos
324             UserList(UserIndex).Invent.Object(slot).Equipped = 0
326             UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
328             UserList(UserIndex).Invent.AnilloEqpSlot = 0
330             UserList(UserIndex).Char.Anillo_Aura = 0
332             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 6))

334             If obj.MagicDamageBonus > 0 Then
336                 Call WriteUpdateDM(UserIndex)
                End If
                
338             If obj.ResistenciaMagica > 0 Then
340                 Call WriteUpdateRM(UserIndex)
                End If
        
        End Select

342     Call UpdateUserInv(False, UserIndex, slot)

        
        Exit Sub

Desequipar_Err:
344     Call RegistrarError(Err.Number, Err.description, "InvUsuario.Desequipar", Erl)
346     Resume Next
        
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
        

100     If ObjData(ObjIndex).Real = 1 Then
102         If Status(UserIndex) = 3 Then
104             FaccionPuedeUsarItem = esArmada(UserIndex)
            Else
106             FaccionPuedeUsarItem = False

            End If

108     ElseIf ObjData(ObjIndex).Caos = 1 Then

110         If Status(UserIndex) = 2 Then
112             FaccionPuedeUsarItem = esCaos(UserIndex)
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
            
160                 If obj.proyectil = 1 Then 'Si es un arco, desequipa el escudo.
            
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
       
232             Case eOBJType.otmagicos
            
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
                        
262                         .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento
                        
264                         If .Stats.UserAtributos(obj.QueAtributo) > MAXATRIBUTOS Then
266                             .Stats.UserAtributos(obj.QueAtributo) = MAXATRIBUTOS
                            End If
                
268                         Call WriteFYA(UserIndex)

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
332                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
            
334                 If Len(obj.CreaGRH) <> 0 Then
336                     .Char.Otra_Aura = obj.CreaGRH
338                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
                    End If
        
                    'Call WriteUpdateExp(UserIndex)
                    'Call CheckUserLevel(UserIndex)
            
340             Case eOBJType.otNUDILLOS
    
342                 If .flags.Muerto = 1 Then
344                     Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
346                 If Not ClasePuedeUsarItem(UserIndex, ObjIndex, slot) Then
348                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                 
350                 If .Invent.WeaponEqpObjIndex > 0 Then
352                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)

                    End If

354                 If .Invent.Object(slot).Equipped Then
356                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
358                 If .Invent.NudilloObjIndex > 0 Then
360                     Call Desequipar(UserIndex, .Invent.NudilloSlot)

                    End If
        
362                 .Invent.Object(slot).Equipped = 1
364                 .Invent.NudilloObjIndex = .Invent.Object(slot).ObjIndex
366                 .Invent.NudilloSlot = slot
        
                    'Falta enviar anim
368                 If .flags.Montado = 0 Then
                
370                     If .flags.Navegando = 0 Then
372                         .Char.WeaponAnim = obj.WeaponAnim
374                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If
            
376                 If obj.SndAura = 0 Then
378                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    Else
380                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                    End If
                 
382                 If Len(obj.CreaGRH) <> 0 Then
384                     .Char.Arma_Aura = obj.CreaGRH
386                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                    End If
    
388             Case eOBJType.otFlechas

390                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Or Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
392                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
394                 If .Invent.Object(slot).Equipped Then
                        'Quitamos del inv el item
396                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
398                 If .Invent.MunicionEqpObjIndex > 0 Then
400                     Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                    End If
        
402                 .Invent.Object(slot).Equipped = 1
404                 .Invent.MunicionEqpObjIndex = .Invent.Object(slot).ObjIndex
406                 .Invent.MunicionEqpSlot = slot

408             Case eOBJType.otArmadura
                
410                 If obj.Ropaje = 0 Then
412                     Call WriteConsoleMsg(UserIndex, "Hay un error con este objeto. Infórmale a un administrador.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Nos aseguramos que puede usarla
414                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Or _
                       Not SexoPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Or _
                       Not CheckRazaUsaRopa(UserIndex, .Invent.Object(slot).ObjIndex) Or _
                       Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
                    
416                     Call WriteConsoleMsg(UserIndex, "Tu clase, género, raza o facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
418                 If .Invent.Object(slot).Equipped Then
                    
420                     Call Desequipar(UserIndex, slot)

422                     If .flags.Navegando = 0 Then
                        
424                         If .flags.Montado = 0 Then
426                             Call DarCuerpoDesnudo(UserIndex)
428                             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                            End If

                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
430                 If .Invent.ArmourEqpObjIndex > 0 Then
432                     errordesc = "Armadura 2"
434                     Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
436                     errordesc = "Armadura 3"

                    End If
  
                    'Lo equipa
438                 If Len(obj.CreaGRH) <> 0 Then
440                     .Char.Body_Aura = obj.CreaGRH
442                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))

                    End If
            
444                 .Invent.Object(slot).Equipped = 1
446                 .Invent.ArmourEqpObjIndex = .Invent.Object(slot).ObjIndex
448                 .Invent.ArmourEqpSlot = slot
                            
450                 If .flags.Montado = 0 Then
                
452                     If .flags.Navegando = 0 Then
                        
454                         .Char.Body = obj.Ropaje
                
456                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
458                         .flags.Desnudo = 0
            
                        End If

                    End If
                
460                 If obj.ResistenciaMagica > 0 Then
462                     Call WriteUpdateRM(UserIndex)
                    End If
    
464             Case eOBJType.otCASCO
                
466                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
468                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
470                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
472                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    
                    End If
                
                    'Si esta equipado lo quita
474                 If .Invent.Object(slot).Equipped Then
476                     Call Desequipar(UserIndex, slot)
                
478                     .Char.CascoAnim = NingunCasco
480                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Exit Sub

                    End If
    
                    'Quita el anterior
482                 If .Invent.CascoEqpObjIndex > 0 Then
484                     Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If
            
486                 errordesc = "Casco"

                    'Lo equipa
488                 If Len(obj.CreaGRH) <> 0 Then
490                     .Char.Head_Aura = obj.CreaGRH
492                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
                    End If
            
494                 .Invent.Object(slot).Equipped = 1
496                 .Invent.CascoEqpObjIndex = .Invent.Object(slot).ObjIndex
498                 .Invent.CascoEqpSlot = slot
            
500                 If .flags.Navegando = 0 Then
502                     .Char.CascoAnim = obj.CascoAnim
504                     Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                
506                 If obj.ResistenciaMagica > 0 Then
508                     Call WriteUpdateRM(UserIndex)
                    End If

510             Case eOBJType.otESCUDO

512                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
514                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
516                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
518                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
520                 If .Invent.Object(slot).Equipped Then
522                     Call Desequipar(UserIndex, slot)
                 
524                     .Char.ShieldAnim = NingunEscudo

526                     If .flags.Montado = 0 Then
528                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                        Exit Sub

                    End If
     
                    'Quita el anterior
530                 If .Invent.EscudoEqpObjIndex > 0 Then
532                     Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
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
540                             Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
542                             Call WriteConsoleMsg(UserIndex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                            End If
                        End If

                    End If
            
544                 errordesc = "Escudo"
             
546                 If Len(obj.CreaGRH) <> 0 Then
548                     .Char.Escudo_Aura = obj.CreaGRH
550                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
                    End If

552                 .Invent.Object(slot).Equipped = 1
554                 .Invent.EscudoEqpObjIndex = .Invent.Object(slot).ObjIndex
556                 .Invent.EscudoEqpSlot = slot
                 
558                 If .flags.Navegando = 0 Then
560                     If .flags.Montado = 0 Then
562                         .Char.ShieldAnim = obj.ShieldAnim
564                         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                    End If
                
566                 If obj.ResistenciaMagica > 0 Then
568                     Call WriteUpdateRM(UserIndex)
                    End If
                
570             Case eOBJType.otAnillos

572                 If Not ClasePuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex, slot) Then
574                     Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

576                 If Not FaccionPuedeUsarItem(UserIndex, .Invent.Object(slot).ObjIndex) Then
578                     Call WriteConsoleMsg(UserIndex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Si esta equipado lo quita
580                 If .Invent.Object(slot).Equipped Then
582                     Call Desequipar(UserIndex, slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
584                 If .Invent.AnilloEqpSlot > 0 Then
586                     Call Desequipar(UserIndex, .Invent.AnilloEqpSlot)
                    End If
                
588                 .Invent.Object(slot).Equipped = 1
590                 .Invent.AnilloEqpObjIndex = .Invent.Object(slot).ObjIndex
592                 .Invent.AnilloEqpSlot = slot
                
594                 If Len(obj.CreaGRH) <> 0 Then
596                     .Char.Anillo_Aura = obj.CreaGRH
598                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Anillo_Aura, False, 6))
                    End If

600                 If obj.MagicDamageBonus > 0 Then
602                     Call WriteUpdateDM(UserIndex)
                    End If
                
604                 If obj.ResistenciaMagica > 0 Then
606                     Call WriteUpdateRM(UserIndex)
                    End If

            End Select
    
        End With

        'Actualiza
608     Call UpdateUserInv(False, UserIndex, slot)

        Exit Sub
    
ErrHandler:
610     Debug.Print errordesc
612     Call LogError("EquiparInvItem Slot:" & slot & " - Error: " & Err.Number & " - Error Description : " & Err.description & "- " & errordesc)

End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then
102         CheckRazaUsaRopa = True
            Exit Function

        End If

104     Select Case UserList(UserIndex).raza

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

100     If UserList(UserIndex).Invent.Object(slot).Amount = 0 Then Exit Sub

102     obj = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex)

104     If obj.OBJType = eOBJType.otWeapon Then
106         If obj.proyectil = 1 Then

                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
108             If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
            Else

                'dagas
110             If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub

            End If

        Else

112         If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
114         If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then Exit Sub

        End If

116     If UserList(UserIndex).flags.Meditando Then
118         UserList(UserIndex).flags.Meditando = False
120         UserList(UserIndex).Char.FX = 0
122         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
        End If

124     If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
126         Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

128     If UserList(UserIndex).Stats.ELV < obj.MinELV Then
130         Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

132     ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
134     UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
136     UserList(UserIndex).flags.TargetObjInvSlot = slot

138     Select Case obj.OBJType

            Case eOBJType.otUseOnce

140             If UserList(UserIndex).flags.Muerto = 1 Then
142                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                'Usa el item
144             UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + obj.MinHam

146             If UserList(UserIndex).Stats.MinHam > UserList(UserIndex).Stats.MaxHam Then UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
148             UserList(UserIndex).flags.Hambre = 0
150             Call WriteUpdateHungerAndThirst(UserIndex)
                'Sonido
        
152             If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
154                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                Else
156                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                End If
        
                'Quitamos del inv el item
158             Call QuitarUserInvItem(UserIndex, slot, 1)
        
160             Call UpdateUserInv(False, UserIndex, slot)

162         Case eOBJType.otGuita

164             If UserList(UserIndex).flags.Muerto = 1 Then
166                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
168             UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(slot).Amount
170             UserList(UserIndex).Invent.Object(slot).Amount = 0
172             UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
174             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
176             Call UpdateUserInv(False, UserIndex, slot)
178             Call WriteUpdateGold(UserIndex)
        
180         Case eOBJType.otWeapon

182             If UserList(UserIndex).flags.Muerto = 1 Then
184                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
186             If Not UserList(UserIndex).Stats.MinSta > 0 Then
188                 Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
190             If ObjData(ObjIndex).proyectil = 1 Then
                    'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
192                 Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                Else

194                 If UserList(UserIndex).flags.TargetObj = Leña Then
196                     If UserList(UserIndex).Invent.Object(slot).ObjIndex = DAGA Then
198                         Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)

                        End If

                    End If

                End If
        
                'REVISAR LADDER
                'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
200             If UserList(UserIndex).Invent.Object(slot).Equipped = 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
202         Case eOBJType.otHerramientas

204             If UserList(UserIndex).flags.Muerto = 1 Then
206                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
208             If Not UserList(UserIndex).Stats.MinSta > 0 Then
210                 Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
212             If UserList(UserIndex).Invent.Object(slot).Equipped = 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
214                 Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

216             Select Case obj.Subtipo
                
                    Case 1, 2  ' Herramientas del Pescador - Caña y Red
218                     Call WriteWorkRequestTarget(UserIndex, eSkill.Pescar)
                
220                 Case 3     ' Herramientas de Alquimia - Tijeras
222                     Call WriteWorkRequestTarget(UserIndex, eSkill.Alquimia)
                
224                 Case 4     ' Herramientas de Alquimia - Olla
226                     Call EnivarObjConstruiblesAlquimia(UserIndex)
228                     Call WriteShowAlquimiaForm(UserIndex)
                
230                 Case 5     ' Herramientas de Carpinteria - Serrucho
232                     Call EnivarObjConstruibles(UserIndex)
234                     Call WriteShowCarpenterForm(UserIndex)
                
236                 Case 6     ' Herramientas de Tala - Hacha
238                     Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)

240                 Case 7     ' Herramientas de Herrero - Martillo
242                     Call WriteConsoleMsg(UserIndex, "Debes hacer click derecho sobre el yunque.", FontTypeNames.FONTTYPE_INFOIAO)

244                 Case 8     ' Herramientas de Mineria - Piquete
246                     Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
                
248                 Case 9     ' Herramientas de Sastreria - Costurero
250                     Call EnivarObjConstruiblesSastre(UserIndex)
252                     Call WriteShowSastreForm(UserIndex)

                End Select
    
254         Case eOBJType.otPociones

256             If UserList(UserIndex).flags.Muerto = 1 Then
258                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
260             UserList(UserIndex).flags.TomoPocion = True
262             UserList(UserIndex).flags.TipoPocion = obj.TipoPocion
                
                Dim CabezaFinal  As Integer

                Dim CabezaActual As Integer

264             Select Case UserList(UserIndex).flags.TipoPocion
        
                    Case 1 'Modif la agilidad
266                     UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                        'Usa el item
268                     UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                
270                     If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                    
272                     If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
                
274                     Call WriteFYA(UserIndex)
                
                        'Quitamos del inv el item
276                     Call QuitarUserInvItem(UserIndex, slot, 1)

278                     If obj.Snd1 <> 0 Then
280                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        Else
282                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If
        
284                 Case 2 'Modif la fuerza
286                     UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                        'Usa el item
288                     UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                
290                     If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                
292                     If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
                
                        'Quitamos del inv el item
294                     Call QuitarUserInvItem(UserIndex, slot, 1)

296                     If obj.Snd1 <> 0 Then
298                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        Else
300                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If

302                     Call WriteFYA(UserIndex)

304                 Case 3 'Pocion roja, restaura HP
                
                        'Usa el item
306                     UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)

308                     If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                
                        'Quitamos del inv el item
310                     Call QuitarUserInvItem(UserIndex, slot, 1)

312                     If obj.Snd1 <> 0 Then
314                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
                        Else
316                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If
            
318                 Case 4 'Pocion azul, restaura MANA
            
                        Dim porcentajeRec As Byte
320                     porcentajeRec = obj.Porcentaje
                
                        'Usa el item
322                     UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, porcentajeRec)

324                     If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                
                        'Quitamos del inv el item
326                     Call QuitarUserInvItem(UserIndex, slot, 1)

328                     If obj.Snd1 <> 0 Then
330                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
                        Else
332                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If
                
334                 Case 5 ' Pocion violeta

336                     If UserList(UserIndex).flags.Envenenado > 0 Then
338                         UserList(UserIndex).flags.Envenenado = 0
340                         Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                            'Quitamos del inv el item
342                         Call QuitarUserInvItem(UserIndex, slot, 1)

344                         If obj.Snd1 <> 0 Then
346                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
                            Else
348                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                            End If

                        Else
350                         Call WriteConsoleMsg(UserIndex, "¡No te encuentras envenenado!", FontTypeNames.FONTTYPE_INFO)

                        End If
                
352                 Case 6  ' Remueve Parálisis

354                     If UserList(UserIndex).flags.Paralizado = 1 Or UserList(UserIndex).flags.Inmovilizado = 1 Then
356                         If UserList(UserIndex).flags.Paralizado = 1 Then
358                             UserList(UserIndex).flags.Paralizado = 0
360                             Call WriteParalizeOK(UserIndex)

                            End If
                        
362                         If UserList(UserIndex).flags.Inmovilizado = 1 Then
364                             UserList(UserIndex).Counters.Inmovilizado = 0
366                             UserList(UserIndex).flags.Inmovilizado = 0
368                             Call WriteInmovilizaOK(UserIndex)

                            End If
                        
                        
                        
370                         Call QuitarUserInvItem(UserIndex, slot, 1)

372                         If obj.Snd1 <> 0 Then
374                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
                            Else
376                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(255, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                            End If

378                         Call WriteConsoleMsg(UserIndex, "Te has removido la paralizis.", FontTypeNames.FONTTYPE_INFOIAO)
                        Else
380                         Call WriteConsoleMsg(UserIndex, "No estas paralizado.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
382                 Case 7  ' Pocion Naranja
384                     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)

386                     If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
                    
                        'Quitamos del inv el item
388                     Call QuitarUserInvItem(UserIndex, slot, 1)

390                     If obj.Snd1 <> 0 Then
392                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            
                        Else
394                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If

396                 Case 8  ' Pocion cambio cara

398                     Select Case UserList(UserIndex).genero

                            Case eGenero.Hombre

400                             Select Case UserList(UserIndex).raza

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

426                             Select Case UserList(UserIndex).raza

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
            
450                     UserList(UserIndex).Char.Head = CabezaFinal
452                     UserList(UserIndex).OrigChar.Head = CabezaFinal
454                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, CabezaFinal, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        'Quitamos del inv el item
456                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 102, 0))

458                     If CabezaActual <> CabezaFinal Then
460                         Call QuitarUserInvItem(UserIndex, slot, 1)
                        Else
462                         Call WriteConsoleMsg(UserIndex, "¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

464                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    
466                 Case 9  ' Pocion sexo
    
468                     Select Case UserList(UserIndex).genero

                            Case eGenero.Hombre
470                             UserList(UserIndex).genero = eGenero.Mujer
                    
472                         Case eGenero.Mujer
474                             UserList(UserIndex).genero = eGenero.Hombre
                    
                        End Select
            
476                     Select Case UserList(UserIndex).genero

                            Case eGenero.Hombre

478                             Select Case UserList(UserIndex).raza

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

504                             Select Case UserList(UserIndex).raza

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
            
528                     UserList(UserIndex).Char.Head = CabezaFinal
530                     UserList(UserIndex).OrigChar.Head = CabezaFinal
532                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, CabezaFinal, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                        'Quitamos del inv el item
534                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 102, 0))
536                     Call QuitarUserInvItem(UserIndex, slot, 1)

538                     If obj.Snd1 <> 0 Then
540                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        Else
542                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If
                
544                 Case 10  ' Invisibilidad
            
546                     If UserList(UserIndex).flags.invisible = 0 Then
548                         UserList(UserIndex).flags.invisible = 1
550                         UserList(UserIndex).Counters.Invisibilidad = obj.DuracionEfecto
552                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
554                         Call WriteContadores(UserIndex)
556                         Call QuitarUserInvItem(UserIndex, slot, 1)

558                         If obj.Snd1 <> 0 Then
560                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            
                            Else
562                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("123", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                            End If

564                         Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras...", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        
                        Else
566                         Call WriteConsoleMsg(UserIndex, "Ya estas invisible.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            Exit Sub

                        End If
                    
568                 Case 11  ' Experiencia

                        Dim HR   As Integer

                        Dim MS   As Integer

                        Dim SS   As Integer

                        Dim secs As Integer

570                     If UserList(UserIndex).flags.ScrollExp = 1 Then
572                         UserList(UserIndex).flags.ScrollExp = obj.CuantoAumento
574                         UserList(UserIndex).Counters.ScrollExperiencia = obj.DuracionEfecto
576                         Call QuitarUserInvItem(UserIndex, slot, 1)
                        
578                         secs = obj.DuracionEfecto
580                         HR = secs \ 3600
582                         MS = (secs Mod 3600) \ 60
584                         SS = (secs Mod 3600) Mod 60

586                         If SS > 9 Then
588                             Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                            Else
590                             Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)

                            End If

                        Else
592                         Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                            Exit Sub

                        End If

594                     Call WriteContadores(UserIndex)

596                     If obj.Snd1 <> 0 Then
598                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        
                        Else
600                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If

602                 Case 12  ' Oro
            
604                     If UserList(UserIndex).flags.ScrollOro = 1 Then
606                         UserList(UserIndex).flags.ScrollOro = obj.CuantoAumento
608                         UserList(UserIndex).Counters.ScrollOro = obj.DuracionEfecto
610                         Call QuitarUserInvItem(UserIndex, slot, 1)
612                         secs = obj.DuracionEfecto
614                         HR = secs \ 3600
616                         MS = (secs Mod 3600) \ 60
618                         SS = (secs Mod 3600) Mod 60

620                         If SS > 9 Then
622                             Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                            Else
624                             Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)

                            End If
                        
                        Else
626                         Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                            Exit Sub

                        End If

628                     Call WriteContadores(UserIndex)

630                     If obj.Snd1 <> 0 Then
632                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        
                        Else
634                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If

636                 Case 13
                
638                     Call QuitarUserInvItem(UserIndex, slot, 1)
640                     UserList(UserIndex).flags.Envenenado = 0
642                     UserList(UserIndex).flags.Incinerado = 0
                    
644                     If UserList(UserIndex).flags.Inmovilizado = 1 Then
646                         UserList(UserIndex).Counters.Inmovilizado = 0
648                         UserList(UserIndex).flags.Inmovilizado = 0
650                         Call WriteInmovilizaOK(UserIndex)
                        

                        End If
                    
652                     If UserList(UserIndex).flags.Paralizado = 1 Then
654                         UserList(UserIndex).flags.Paralizado = 0
656                         Call WriteParalizeOK(UserIndex)
                        

                        End If
                    
658                     If UserList(UserIndex).flags.Ceguera = 1 Then
660                         UserList(UserIndex).flags.Ceguera = 0
662                         Call WriteBlindNoMore(UserIndex)
                        

                        End If
                    
664                     If UserList(UserIndex).flags.Maldicion = 1 Then
666                         UserList(UserIndex).flags.Maldicion = 0
668                         UserList(UserIndex).Counters.Maldicion = 0

                        End If
                    
670                     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
672                     UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
674                     UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
676                     UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
678                     UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
                    
680                     UserList(UserIndex).flags.Hambre = 0
682                     UserList(UserIndex).flags.Sed = 0
                    
684                     Call WriteUpdateHungerAndThirst(UserIndex)
686                     Call WriteConsoleMsg(UserIndex, "Donador> Te sentis sano y lleno.", FontTypeNames.FONTTYPE_WARNING)

688                     If obj.Snd1 <> 0 Then
690                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        
                        Else
692                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If

694                 Case 14
                
696                     If UserList(UserIndex).flags.BattleModo = 1 Then
698                         Call WriteConsoleMsg(UserIndex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub

                        End If
                    
700                     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = CARCEL Then
702                         Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                    
                        Dim Map     As Integer

                        Dim X       As Byte

                        Dim Y       As Byte

                        Dim DeDonde As WorldPos

704                     Call QuitarUserInvItem(UserIndex, slot, 1)
            
706                     Select Case UserList(UserIndex).Hogar

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
                    
740                     Call FindLegalPos(UserIndex, Map, X, Y)
742                     Call WarpUserChar(UserIndex, Map, X, Y, True)
744                     Call WriteConsoleMsg(UserIndex, "Ya estas a salvo...", FontTypeNames.FONTTYPE_WARNING)

746                     If obj.Snd1 <> 0 Then
748                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        
                        Else
750                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If

752                 Case 15  ' Aliento de sirena
                        
754                     If UserList(UserIndex).Counters.Oxigeno >= 3540 Then
                        
756                         Call WriteConsoleMsg(UserIndex, "No podes acumular más de 59 minutos de oxigeno.", FontTypeNames.FONTTYPE_INFOIAO)
758                         secs = UserList(UserIndex).Counters.Oxigeno
760                         HR = secs \ 3600
762                         MS = (secs Mod 3600) \ 60
764                         SS = (secs Mod 3600) Mod 60

766                         If SS > 9 Then
768                             Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                            Else
770                             Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)

                            End If

                        Else
                            
772                         UserList(UserIndex).Counters.Oxigeno = UserList(UserIndex).Counters.Oxigeno + obj.DuracionEfecto
774                         Call QuitarUserInvItem(UserIndex, slot, 1)
                            
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
                            
776                         UserList(UserIndex).flags.Ahogandose = 0
778                         Call WriteOxigeno(UserIndex)
                            
780                         Call WriteContadores(UserIndex)

782                         If obj.Snd1 <> 0 Then
784                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            
                            Else
786                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                            End If

                        End If

788                 Case 16 ' Divorcio

790                     If UserList(UserIndex).flags.Casado = 1 Then

                            Dim tUser As Integer

                            'UserList(UserIndex).flags.Pareja
792                         tUser = NameIndex(UserList(UserIndex).flags.Pareja)
794                         Call QuitarUserInvItem(UserIndex, slot, 1)
                        
796                         If tUser <= 0 Then

                                Dim FileUser As String

798                             FileUser = CharPath & UCase$(UserList(UserIndex).flags.Pareja) & ".chr"
                                'Call WriteVar(FileUser, "FLAGS", "CASADO", 0)
                                'Call WriteVar(FileUser, "FLAGS", "PAREJA", "")
800                             UserList(UserIndex).flags.Casado = 0
802                             UserList(UserIndex).flags.Pareja = ""
804                             Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
806                             UserList(UserIndex).MENSAJEINFORMACION = UserList(UserIndex).name & " se ha divorciado de ti."

                            Else
808                             UserList(tUser).flags.Casado = 0
810                             UserList(tUser).flags.Pareja = ""
812                             UserList(UserIndex).flags.Casado = 0
814                             UserList(UserIndex).flags.Pareja = ""
816                             Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
818                             Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " se ha divorciado de ti.", FontTypeNames.FONTTYPE_INFOIAO)
                            
                            End If

820                         If obj.Snd1 <> 0 Then
822                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            
                            Else
824                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                            End If
                    
                        Else
826                         Call WriteConsoleMsg(UserIndex, "No estas casado.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

828                 Case 17 'Cara legendaria

830                     Select Case UserList(UserIndex).genero

                            Case eGenero.Hombre

832                             Select Case UserList(UserIndex).raza

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

858                             Select Case UserList(UserIndex).raza

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

882                     CabezaActual = UserList(UserIndex).OrigChar.Head
                        
884                     UserList(UserIndex).Char.Head = CabezaFinal
886                     UserList(UserIndex).OrigChar.Head = CabezaFinal
888                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, CabezaFinal, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                        'Quitamos del inv el item
890                     If CabezaActual <> CabezaFinal Then
892                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 102, 0))
894                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
896                         Call QuitarUserInvItem(UserIndex, slot, 1)
                        Else
898                         Call WriteConsoleMsg(UserIndex, "¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!", FontTypeNames.FONTTYPE_INFOIAO)

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
914                             UserList(UserIndex).Char.ParticulaFx = Particula
916                             UserList(UserIndex).Char.loops = Tiempo

                            End If
                            
918                         If sobrechar = 1 Then
920                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, Particula, Tiempo))
                            Else
                            
922                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, Particula, Tiempo, False))

                                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, Particula, Tiempo))
                            End If

                        End If
                        
924                     If obj.CreaFX <> 0 Then
926                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                            
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, obj.CreaFX, 0))
                            ' PrepareMessageCreateFX
                        End If
                        
928                     If obj.Snd1 <> 0 Then
930                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                        End If
                        
932                     Call QuitarUserInvItem(UserIndex, slot, 1)

934                 Case 19 ' Reseteo de skill

                        Dim S As Byte
                
936                     If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 80 Then
938                         Call WriteConsoleMsg(UserIndex, "Has fundado un clan, no podes resetar tus skills. ", FontTypeNames.FONTTYPE_INFOIAO)
                            Exit Sub

                        End If
                    
940                     For S = 1 To NUMSKILLS
942                         UserList(UserIndex).Stats.UserSkills(S) = 0
944                     Next S
                    
                        Dim SkillLibres As Integer
                    
946                     SkillLibres = 5
948                     SkillLibres = SkillLibres + (5 * UserList(UserIndex).Stats.ELV)
                     
950                     UserList(UserIndex).Stats.SkillPts = SkillLibres
952                     Call WriteLevelUp(UserIndex, UserList(UserIndex).Stats.SkillPts)
                    
954                     Call WriteConsoleMsg(UserIndex, "Tus skills han sido reseteados.", FontTypeNames.FONTTYPE_INFOIAO)
956                     Call QuitarUserInvItem(UserIndex, slot, 1)

958                 Case 20
                
960                     If UserList(UserIndex).Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
962                         UserList(UserIndex).Stats.InventLevel = UserList(UserIndex).Stats.InventLevel + 1
964                         UserList(UserIndex).CurrentInventorySlots = getMaxInventorySlots(UserIndex)
966                         Call WriteInventoryUnlockSlots(UserIndex)
968                         Call WriteConsoleMsg(UserIndex, "Has aumentado el espacio de tu inventario!", FontTypeNames.FONTTYPE_INFO)
970                         Call QuitarUserInvItem(UserIndex, slot, 1)
                        Else
972                         Call WriteConsoleMsg(UserIndex, "Ya has desbloqueado todos los casilleros disponibles.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                
                End Select

974             Call WriteUpdateUserStats(UserIndex)
976             Call UpdateUserInv(False, UserIndex, slot)

978         Case eOBJType.otBebidas

980             If UserList(UserIndex).flags.Muerto = 1 Then
982                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

984             UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed

986             If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
988             UserList(UserIndex).flags.Sed = 0
990             Call WriteUpdateHungerAndThirst(UserIndex)
        
                'Quitamos del inv el item
992             Call QuitarUserInvItem(UserIndex, slot, 1)
        
994             If obj.Snd1 <> 0 Then
996                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
                Else
998                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                 End If
        
1000             Call UpdateUserInv(False, UserIndex, slot)
        
1002         Case eOBJType.OtCofre

1004             If UserList(UserIndex).flags.Muerto = 1 Then
1006                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

                 'Quitamos del inv el item
1008             Call QuitarUserInvItem(UserIndex, slot, 1)
1010             Call UpdateUserInv(False, UserIndex, slot)
        
1012             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha abierto un " & obj.name & " y obtuvo...", FontTypeNames.FONTTYPE_New_DONADOR))
        
1014             If obj.Snd1 <> 0 Then
1016                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

                 End If
        
1018             If obj.CreaFX <> 0 Then
1020                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, obj.CreaFX, 0))

                 End If
        
                 Dim i As Byte

1022             If obj.Subtipo = 1 Then

1024                 For i = 1 To obj.CantItem

1026                     If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, obj.Item(i))
1028                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
1030                 Next i
        
                 Else
        
1032                 For i = 1 To obj.CantEntrega

                         Dim indexobj As Byte
                
1034                     indexobj = RandomNumber(1, obj.CantItem)
            
                         Dim index As obj

1036                     index.ObjIndex = obj.Item(indexobj).ObjIndex
1038                     index.Amount = obj.Item(indexobj).Amount

1040                     If Not MeterItemEnInventario(UserIndex, index) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, index)
1042                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(index.ObjIndex).name & " (" & index.Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
1044                 Next i

                 End If
    
1046         Case eOBJType.otLlaves
1048             Call WriteConsoleMsg(UserIndex, "Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.", FontTypeNames.FONTTYPE_INFO)
    
1050         Case eOBJType.otBotellaVacia

1052             If UserList(UserIndex).flags.Muerto = 1 Then
1054                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1056             If (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
1058                 Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1060             MiObj.Amount = 1
1062             MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).IndexAbierta
1064             Call QuitarUserInvItem(UserIndex, slot, 1)

1066             If Not MeterItemEnInventario(UserIndex, MiObj) Then
1068                 Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

                 End If
        
1070             Call UpdateUserInv(False, UserIndex, slot)
    
1072         Case eOBJType.otBotellaLlena

1074             If UserList(UserIndex).flags.Muerto = 1 Then
1076                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1078             UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed

1080             If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
1082             UserList(UserIndex).flags.Sed = 0
1084             Call WriteUpdateHungerAndThirst(UserIndex)
1086             MiObj.Amount = 1
1088             MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).IndexCerrada
1090             Call QuitarUserInvItem(UserIndex, slot, 1)

1092             If Not MeterItemEnInventario(UserIndex, MiObj) Then
1094                 Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

                 End If
        
1096             Call UpdateUserInv(False, UserIndex, slot)
    
1098         Case eOBJType.otPergaminos

1100             If UserList(UserIndex).flags.Muerto = 1 Then
1102                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
                 'Call LogError(UserList(UserIndex).Name & " intento aprender el hechizo " & ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex)
        
1104             If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex, slot) Then

                     'If UserList(UserIndex).Stats.MaxMAN > 0 Then
1106                 If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then
1108                     Call AgregarHechizo(UserIndex, slot)
1110                     Call UpdateUserInv(False, UserIndex, slot)
                         ' Call LogError(UserList(UserIndex).Name & " lo aprendio.")
                     Else
1112                     Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                     End If

                     ' Else
                     '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_WARNING)
                     'End If
                 Else
             
1114                 Call WriteConsoleMsg(UserIndex, "Por mas que lo intentas, no podés comprender el manuescrito.", FontTypeNames.FONTTYPE_INFO)
   
                 End If
        
1116         Case eOBJType.otMinerales

1118             If UserList(UserIndex).flags.Muerto = 1 Then
1120                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1122             Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
1124         Case eOBJType.otInstrumentos

1126             If UserList(UserIndex).flags.Muerto = 1 Then
1128                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1130             If obj.Real Then '¿Es el Cuerno Real?
1132                 If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1134                     If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
1136                         Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If

1138                     Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                         Exit Sub
                     Else
1140                     Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                         Exit Sub

                     End If

1142             ElseIf obj.Caos Then '¿Es el Cuerno Legión?

1144                 If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1146                     If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
1148                         Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If

1150                     Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                         Exit Sub
                     Else
1152                     Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                         Exit Sub

                     End If

                 End If

                 'Si llega aca es porque es o Laud o Tambor o Flauta
1154             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
       
1156         Case eOBJType.otBarcos
            
                 'Verifica si tiene el nivel requerido para navegar, siendo Trabajador o Pirata
1158             If UserList(UserIndex).Stats.ELV < 20 And (UserList(UserIndex).clase = eClass.Trabajador Or UserList(UserIndex).clase = eClass.Pirat) Then
1160                 Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 'Verifica si tiene el nivel requerido para navegar, sin ser Trabajador o Pirata
1162             ElseIf UserList(UserIndex).Stats.ELV < 25 Then
1164                 Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
                 'If obj.Subtipo = 0 Then
1166             If UserList(UserIndex).flags.Navegando = 0 Then
1168                 If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, True, False) Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, True, False) Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, True, False) Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, True, False)) And UserList(UserIndex).flags.Navegando = 0) Or UserList(UserIndex).flags.Navegando = 1 Then
1170                     Call DoNavega(UserIndex, obj, slot)
                     Else
1172                     Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)

                     End If

                 Else 'Ladder 10-02-2010

1174                 If UserList(UserIndex).Invent.BarcoObjIndex <> UserList(UserIndex).Invent.Object(slot).ObjIndex Then
1176                     Call DoReNavega(UserIndex, obj, slot)
                     Else

1178                     If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X - 1, UserList(UserIndex).Pos.Y, False, True) Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y - 1, False, True) Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X + 1, UserList(UserIndex).Pos.Y, False, True) Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y + 1, False, True)) And UserList(UserIndex).flags.Navegando = 1) Or UserList(UserIndex).flags.Navegando = 0 Then
1180                         Call DoNavega(UserIndex, obj, slot)
                         Else
1182                         Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte a la costa para dejar la barca!", FontTypeNames.FONTTYPE_INFO)

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
    
1186             If UserList(UserIndex).flags.Muerto = 1 Then
1188                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡Estas muerto! Los fantasmas no pueden montar.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
            
1190             If UserList(UserIndex).flags.Navegando = 1 Then
1192                 Call WriteConsoleMsg(UserIndex, "Debes dejar de navegar para poder montarté.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If

1194             If MapInfo(UserList(UserIndex).Pos.Map).zone = "DUNGEON" Then
1196                 Call WriteConsoleMsg(UserIndex, "No podes cabalgar dentro de un dungeon.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1198             Call DoMontar(UserIndex, obj, slot)

1200         Case eOBJType.OtDonador

1202             Select Case obj.Subtipo

                     Case 1
            
1204                     If UserList(UserIndex).Counters.Pena <> 0 Then
1206                         Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If
                
1208                     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = CARCEL Then
1210                         Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                             Exit Sub

                         End If
            
1212                     Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1214                     Call WriteConsoleMsg(UserIndex, "Has viajado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
1216                     Call QuitarUserInvItem(UserIndex, slot, 1)
1218                     Call UpdateUserInv(False, UserIndex, slot)
                
1220                 Case 2

1222                     If DonadorCheck(UserList(UserIndex).Cuenta) = 0 Then
1224                         Call DonadorTiempo(UserList(UserIndex).Cuenta, CLng(obj.CuantoAumento))
1226                         Call WriteConsoleMsg(UserIndex, "Donación> Se han agregado " & obj.CuantoAumento & " dias de donador a tu cuenta. Relogea tu personaje para empezar a disfrutar la experiencia.", FontTypeNames.FONTTYPE_WARNING)
1228                         Call QuitarUserInvItem(UserIndex, slot, 1)
1230                         Call UpdateUserInv(False, UserIndex, slot)
                         Else
1232                         Call DonadorTiempo(UserList(UserIndex).Cuenta, CLng(obj.CuantoAumento))
1234                         Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & CLng(obj.CuantoAumento) & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
1236                         UserList(UserIndex).donador.activo = 1
1238                         Call QuitarUserInvItem(UserIndex, slot, 1)
1240                         Call UpdateUserInv(False, UserIndex, slot)

                             'Call WriteConsoleMsg(UserIndex, "Donación> Debes esperar a que finalice el periodo existente para renovar tu suscripción.", FontTypeNames.FONTTYPE_INFOIAO)
                         End If

1242                 Case 3
1244                     Call AgregarCreditosDonador(UserList(UserIndex).Cuenta, CLng(obj.CuantoAumento))
1246                     Call WriteConsoleMsg(UserIndex, "Donación> Tu credito ahora es de " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
1248                     Call QuitarUserInvItem(UserIndex, slot, 1)
1250                     Call UpdateUserInv(False, UserIndex, slot)

                 End Select
     
1252         Case eOBJType.otpasajes

1254             If UserList(UserIndex).flags.Muerto = 1 Then
1256                 Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1258             If UserList(UserIndex).flags.TargetNpcTipo <> Pirata Then
1260                 Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1262             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
1264                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                     'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1266             If UserList(UserIndex).Pos.Map <> obj.DesdeMap Then
                     Rem  Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aquí! Largate!", FontTypeNames.FONTTYPE_INFO)
1268                 Call WriteChatOverHead(UserIndex, "El pasaje no lo compraste aquí! Largate!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                     Exit Sub

                 End If
        
1270             If Not MapaValido(obj.HastaMap) Then
                     Rem Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
1272                 Call WriteChatOverHead(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                     Exit Sub

                 End If

1274             If obj.NecesitaNave > 0 Then
1276                 If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                         Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
1278                     Call WriteChatOverHead(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                         Exit Sub

                     End If

                 End If
            
1280             Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1282             Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
1284             UserList(UserIndex).Stats.MinAGU = 0
1286             UserList(UserIndex).Stats.MinHam = 0
1288             UserList(UserIndex).flags.Sed = 1
1290             UserList(UserIndex).flags.Hambre = 1
1292             Call WriteUpdateHungerAndThirst(UserIndex)
1294             Call QuitarUserInvItem(UserIndex, slot, 1)
1296             Call UpdateUserInv(False, UserIndex, slot)
        
1298         Case eOBJType.otRunas
    
1300             If UserList(UserIndex).Counters.Pena <> 0 Then
1302                 Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1304             If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = CARCEL Then
1306                 Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1308             If UserList(UserIndex).flags.BattleModo = 1 Then
1310                 Call WriteConsoleMsg(UserIndex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                     Exit Sub

                 End If
        
1312             If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 And UserList(UserIndex).flags.Muerto = 0 Then
1314                 Call WriteConsoleMsg(UserIndex, "Solo podes usar tu runa en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub

                 End If
        
1316             If UserList(UserIndex).Accion.AccionPendiente Then
                     Exit Sub

                 End If
        
1318             Select Case ObjData(ObjIndex).TipoRuna
        
                     Case 1, 2

1320                     If UserList(UserIndex).donador.activo = 0 Then ' Donador no espera tiempo
1322                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
1324                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 350, Accion_Barra.Runa))
                         Else
1326                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
1328                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 100, Accion_Barra.Runa))

                         End If

1330                     UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
1332                     UserList(UserIndex).Accion.AccionPendiente = True
1334                     UserList(UserIndex).Accion.TipoAccion = Accion_Barra.Runa
1336                     UserList(UserIndex).Accion.RunaObj = ObjIndex
1338                     UserList(UserIndex).Accion.ObjSlot = slot
            
1340                 Case 3
        
                         Dim parejaindex As Integer

1342                     If Not UserList(UserIndex).flags.BattleModo Then
                
                             'If UserList(UserIndex).donador.activo = 1 Then
1344                         If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
1346                             If UserList(UserIndex).flags.Casado = 1 Then
1348                                 parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                        
1350                                 If parejaindex > 0 Then
1352                                     If UserList(parejaindex).flags.BattleModo = 0 Then
1354                                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 600, False))
1356                                         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.GoToPareja))
1358                                         UserList(UserIndex).Accion.AccionPendiente = True
1360                                         UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
1362                                         UserList(UserIndex).Accion.TipoAccion = Accion_Barra.GoToPareja
                                         Else
1364                                         Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                         End If
                                
                                     Else
1366                                     Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                                     End If

                                 Else
1368                                 Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                                 End If

                             Else
1370                             Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                             End If
                
                             ' Else
                             '  Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                             ' End If
                         Else
1372                         Call WriteConsoleMsg(UserIndex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
        
                         End If
    
                 End Select
        
1374         Case eOBJType.otmapa
1376             Call WriteShowFrmMapa(UserIndex)
        
         End Select

         Exit Sub

hErr:
1378     LogError "Error en useinvitem Usuario: " & UserList(UserIndex).name & " item:" & obj.name & " index: " & UserList(UserIndex).Invent.Object(slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmasConstruibles_Err
        

100     Call WriteBlacksmithWeapons(UserIndex)

        
        Exit Sub

EnivarArmasConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarArmasConstruibles", Erl)
104     Resume Next
        
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruibles_Err
        

100     Call WriteCarpenterObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruibles", Erl)
104     Resume Next
        
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesAlquimia_Err
        

100     Call WriteAlquimistaObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesAlquimia_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)
104     Resume Next
        
End Sub

Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesSastre_Err
        

100     Call WriteSastreObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesSastre_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)
104     Resume Next
        
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmadurasConstruibles_Err
        

100     Call WriteBlacksmithArmors(UserIndex)

        
        Exit Sub

EnivarArmadurasConstruibles_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarArmadurasConstruibles", Erl)
104     Resume Next
        
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)

        On Error Resume Next

100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
102     If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub

104     Call TirarTodosLosItems(UserIndex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
        
        On Error GoTo ItemSeCae_Err
        

100     ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And ObjData(index).OBJType <> eOBJType.otLlaves And ObjData(index).OBJType <> eOBJType.otBarcos And ObjData(index).OBJType <> eOBJType.otMonturas And ObjData(index).NoSeCae = 0 And Not ObjData(index).Intirable = 1 And Not ObjData(index).Destruye = 1 And ObjData(index).donador = 0 And Not ObjData(index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
102     Call RegistrarError(Err.Number, Err.description, "InvUsuario.ItemSeCae", Erl)
104     Resume Next
        
End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal slot As Byte)

100     With UserList(UserIndex)
    
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

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
        
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
    
100     With UserList(UserIndex)
    
102         For i = 1 To .CurrentInventorySlots
    
104             ItemIndex = .Invent.Object(i).ObjIndex

106             If ItemIndex > 0 Then

108                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) Then
110                     NuevaPos.X = 0
112                     NuevaPos.Y = 0
                
114                     If .flags.CarroMineria = 1 Then
                
116                         If ItemIndex = ORO_MINA Or ItemIndex = PLATA_MINA Or ItemIndex = HIERRO_MINA Then
                       
118                             MiObj.Amount = .Invent.Object(i).Amount * 0.3
120                             MiObj.ObjIndex = ItemIndex
                        
122                             Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                    
124                             If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
126                                 Call DropObj(UserIndex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                                End If

                            End If
                    
                        Else
                    
128                         MiObj.Amount = .Invent.Object(i).Amount
130                         MiObj.ObjIndex = ItemIndex
                        
132                         Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                
134                         If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
136                             Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
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

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
        
        On Error GoTo TirarTodosLosItemsNoNewbies_Err
        

        Dim i         As Byte

        Dim NuevaPos  As WorldPos

        Dim MiObj     As obj

        Dim ItemIndex As Integer

100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

102     For i = 1 To UserList(UserIndex).CurrentInventorySlots
104         ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex

106         If ItemIndex > 0 Then
108             If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
110                 NuevaPos.X = 0
112                 NuevaPos.Y = 0
            
                    'Creo MiObj
114                 MiObj.Amount = UserList(UserIndex).Invent.Object(i).ObjIndex
116                 MiObj.ObjIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
118                 Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True

120                 If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
122                     If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

                    End If

                End If

            End If

124     Next i

        
        Exit Sub

TirarTodosLosItemsNoNewbies_Err:
126     Call RegistrarError(Err.Number, Err.description, "InvUsuario.TirarTodosLosItemsNoNewbies", Erl)
128     Resume Next
        
End Sub
