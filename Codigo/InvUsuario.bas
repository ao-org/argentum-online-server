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

    For i = 1 To UserList(Userindex).CurrentInventorySlots
        ObjIndex = UserList(Userindex).Invent.Object(i).ObjIndex

        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And ObjData(ObjIndex).OBJType <> eOBJType.OtDonador And ObjData(ObjIndex).OBJType <> eOBJType.otRunas) Then
                TieneObjetosRobables = True
                Exit Function

            End If
    
        End If

    Next i

End Function

Function ClasePuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer, Optional slot As Byte) As Boolean

    On Error GoTo manejador

    'Call LogTarea("ClasePuedeUsarItem")

    Dim flag As Boolean

    If slot <> 0 Then
        If UserList(Userindex).Invent.Object(slot).Equipped Then
            ClasePuedeUsarItem = True
            Exit Function

        End If

    End If

    If EsGM(Userindex) Then
        ClasePuedeUsarItem = True
        Exit Function

    End If

    'Admins can use ANYTHING!
    'If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    'If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
    Dim i As Integer

    For i = 1 To NUMCLASES

        If ObjData(ObjIndex).ClaseProhibida(i) = UserList(Userindex).clase Then
            ClasePuedeUsarItem = False
            Exit Function

        End If

    Next i

    ' End If
    'End If

    ClasePuedeUsarItem = True

    Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")

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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.QuitarNewbieObj", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.LimpiarInventario", Erl)
        Resume Next
        
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal Userindex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
    '***************************************************
    On Error GoTo ErrHandler

    'If Cantidad > 100000 Then Exit Sub
    If UserList(Userindex).flags.BattleModo = 1 Then Exit Sub

    'SI EL Pjta TIENE ORO LO TIRAMOS
    If (Cantidad > 0) And (Cantidad <= UserList(Userindex).Stats.GLD) Then

        Dim i     As Byte

        Dim MiObj As obj

        Dim Logs  As Long

        'info debug
        Dim loops As Integer
        
        Logs = Cantidad

        Dim Extra    As Long

        Dim TeniaOro As Long

        TeniaOro = UserList(Userindex).Stats.GLD

        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000

        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(Userindex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                Cantidad = Cantidad - MiObj.Amount

            End If

            MiObj.ObjIndex = iORO

            Dim AuxPos As WorldPos
            If UserList(Userindex).clase = eClass.Pirat Then
                AuxPos = TirarItemAlPiso(UserList(Userindex).Pos, MiObj, False)
            Else
                AuxPos = TirarItemAlPiso(UserList(Userindex).Pos, MiObj, True)
            End If
            
            If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - MiObj.Amount

            End If
            
            'info debug
            loops = loops + 1

            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub

            End If
            
        Loop
        
        If EsGM(Userindex) Then
            If MiObj.ObjIndex = iORO Then
                Call LogGM(UserList(Userindex).name, "Tiro: " & Logs & " monedas de oro.")
            Else
                Call LogGM(UserList(Userindex).name, "Tiro cantidad:" & Logs & " Objeto:" & ObjData(MiObj.ObjIndex).name)

            End If

        End If
        
        If TeniaOro = UserList(Userindex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - Extra

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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.QuitarUserInvItem", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.UpdateUserInv", Erl)
        Resume Next
        
End Sub

Sub DropObj(ByVal Userindex As Integer, ByVal slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo DropObj_Err
        

        Dim obj As obj

100     If num > 0 Then
      
102         If num > UserList(Userindex).Invent.Object(slot).Amount Then num = UserList(Userindex).Invent.Object(slot).Amount
104         obj.ObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
106         obj.Amount = num

108         If ObjData(obj.ObjIndex).Destruye = 0 Then

                'Check objeto en el suelo
110             If MapData(UserList(Userindex).Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Then
                  
112                 If num + MapData(UserList(Userindex).Pos.Map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
114                     num = MAX_INVENTORY_OBJS - MapData(UserList(Userindex).Pos.Map, X, Y).ObjInfo.Amount

                    End If
                  
116                 Call MakeObj(obj, Map, X, Y)
118                 Call QuitarUserInvItem(Userindex, slot, num)
120                 Call UpdateUserInv(False, Userindex, slot)
                  
122                 If Not UserList(Userindex).flags.Privilegios And PlayerType.user Then Call LogGM(UserList(Userindex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)
                  
                    'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                    'Es un Objeto que tenemos que loguear?
                    ' If ObjData(obj.ObjIndex).Log = 1 Then
                    '    Call LogDesarrollo(UserList(UserIndex).name & " tiró al piso " & obj.Amount & " " & ObjData(obj.ObjIndex).name)
                    '    ElseIf obj.Amount = 1000 Then 'Es mucha cantidad?
                    '   'Si no es de los prohibidos de loguear, lo logueamos.
                    '  If ObjData(obj.ObjIndex).NoLog <> 1 Then
                    '    Call LogDesarrollo(UserList(UserIndex).name & " tiró del piso " & obj.Amount & " " & ObjData(obj.ObjIndex).name)
                    ' End If
                    ' End If
                Else
                    'Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
124                 Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
126             Call QuitarUserInvItem(Userindex, slot, num)
128             Call UpdateUserInv(False, Userindex, slot)

            End If

        End If

        
        Exit Sub

DropObj_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.DropObj", Erl)
        Resume Next
        
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
104         MapData(Map, X, Y).ObjInfo.ObjIndex = 0
106         MapData(Map, X, Y).ObjInfo.Amount = 0
    
            If ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType <> otTeleport Then
                Call QuitarItemLimpieza(Map, X, Y)
            End If
    
108         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

        
        Exit Sub

EraseObj_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.EraseObj", Erl)
        Resume Next
        
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
                If ObjData(obj.ObjIndex).OBJType <> otTeleport Then
                    Call AgregarItemLimpiza(Map, X, Y, MapData(Map, X, Y).ObjInfo.ObjIndex <> 0)
                End If
            
106             MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex

108             If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
110                 MapData(Map, X, Y).ObjInfo.Amount = ObjData(obj.ObjIndex).VidaUtil
                Else
112                 MapData(Map, X, Y).ObjInfo.Amount = obj.Amount

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
114             Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(obj.ObjIndex, X, Y))
                
            End If
    
        End If
        
        Exit Sub

MakeObj_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.MakeObj", Erl)

        Resume Next
        
End Sub

Function MeterItemEnInventario(ByVal Userindex As Integer, ByRef MiObj As obj) As Boolean

    On Error GoTo ErrHandler

    'Call LogTarea("MeterItemEnInventario")
 
    Dim X    As Integer

    Dim Y    As Integer

    Dim slot As Byte

    '¿el user ya tiene un objeto del mismo tipo? ?????
    If MiObj.ObjIndex = 12 Then
        UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + MiObj.Amount

    Else
    
        slot = 1

        Do Until UserList(Userindex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And UserList(Userindex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
            slot = slot + 1

            If slot > UserList(Userindex).CurrentInventorySlots Then
                Exit Do

            End If

        Loop
        
        'Sino busca un slot vacio
        If slot > UserList(Userindex).CurrentInventorySlots Then
            slot = 1

            Do Until UserList(Userindex).Invent.Object(slot).ObjIndex = 0
                slot = slot + 1

                If slot > UserList(Userindex).CurrentInventorySlots Then
                    'Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteLocaleMsg(Userindex, "328", FontTypeNames.FONTTYPE_FIGHT)
                    MeterItemEnInventario = False
                    Exit Function

                End If

            Loop
            UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems + 1

        End If
        
        'Mete el objeto
        If UserList(Userindex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            UserList(Userindex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex
            UserList(Userindex).Invent.Object(slot).Amount = UserList(Userindex).Invent.Object(slot).Amount + MiObj.Amount
        Else
            UserList(Userindex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS

        End If
        
        MeterItemEnInventario = True
           
        Call UpdateUserInv(False, Userindex, slot)

    End If

    WriteUpdateGold (Userindex)
    MeterItemEnInventario = True

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
    
    slot = 1

    Do Until Npclist(NpcIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And Npclist(NpcIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
        slot = slot + 1

        If slot > MAX_INVENTORY_SLOTS Then
            Exit Do

        End If

    Loop
        
    'Sino busca un slot vacio
    If slot > MAX_INVENTORY_SLOTS Then
        slot = 1

        Do Until Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
            slot = slot + 1

            If slot > MAX_INVENTORY_SLOTS Then
                Rem Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
                MeterItemEnInventarioDeNpc = False
                Exit Function

            End If

        Loop
        Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

    End If

    MeterItemEnInventarioDeNpc = True

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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.GetObj", Erl)
        Resume Next
        
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
                
                If obj.MagicDamageBonus > 0 Then
                    Call WriteUpdateDM(Userindex)
                End If
    
134         Case eOBJType.otFlechas
136             UserList(Userindex).Invent.Object(slot).Equipped = 0
138             UserList(Userindex).Invent.MunicionEqpObjIndex = 0
140             UserList(Userindex).Invent.MunicionEqpSlot = 0
    
                ' Case eOBJType.otAnillos
                '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
                '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
                ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
            
142         Case eOBJType.otHerramientas
144             UserList(Userindex).Invent.Object(slot).Equipped = 0
146             UserList(Userindex).Invent.HerramientaEqpObjIndex = 0
148             UserList(Userindex).Invent.HerramientaEqpSlot = 0

150             If UserList(Userindex).flags.UsandoMacro = True Then
152                 Call WriteMacroTrabajoToggle(Userindex, False)
                End If
        
154             UserList(Userindex).Char.WeaponAnim = NingunArma
            
156             If UserList(Userindex).flags.Montado = 0 Then
158                 Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                End If
       
160         Case eOBJType.otmagicos
    
162             Select Case obj.EfectoMagico

                    Case 1 'Regenera Energia
164                     UserList(Userindex).flags.RegeneracionSta = 0

166                 Case 2 'Modifica los Atributos
168                     UserList(Userindex).Stats.UserAtributos(obj.QueAtributo) = UserList(Userindex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                
170                     UserList(Userindex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(Userindex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                        ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
172                     Call WriteFYA(Userindex)

174                 Case 3 'Modifica los skills
176                     UserList(Userindex).Stats.UserSkills(obj.QueSkill) = UserList(Userindex).Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento

178                 Case 4 ' Regeneracion Vida
180                     UserList(Userindex).flags.RegeneracionHP = 0

182                 Case 5 ' Regeneracion Mana
184                     UserList(Userindex).flags.RegeneracionMana = 0

186                 Case 6 'Aumento Golpe
188                     UserList(Userindex).Stats.MaxHit = UserList(Userindex).Stats.MaxHit - obj.CuantoAumento
190                     UserList(Userindex).Stats.MinHIT = UserList(Userindex).Stats.MinHIT - obj.CuantoAumento

192                 Case 7 '
                
194                 Case 9 ' Orbe Ignea
196                     UserList(Userindex).flags.NoMagiaEfeceto = 0

198                 Case 10
200                     UserList(Userindex).flags.incinera = 0

202                 Case 11
204                     UserList(Userindex).flags.Paraliza = 0

206                 Case 12

208                     If UserList(Userindex).flags.Muerto = 0 Then UserList(Userindex).flags.CarroMineria = 0
                
210                 Case 14
212                     'UserList(UserIndex).flags.DañoMagico = 0
                
214                 Case 15 'Pendiete del Sacrificio
216                     UserList(Userindex).flags.PendienteDelSacrificio = 0
                 
218                 Case 16
220                     UserList(Userindex).flags.NoPalabrasMagicas = 0

222                 Case 17 'Sortija de la verdad
224                     UserList(Userindex).flags.NoDetectable = 0

226                 Case 18 ' Pendiente del Experto
228                     UserList(Userindex).flags.PendienteDelExperto = 0

230                 Case 19
232                     UserList(Userindex).flags.Envenena = 0

234                 Case 20 ' anillo de las sombras
236                     UserList(Userindex).flags.AnilloOcultismo = 0
                
                End Select
        
244             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 5))
246             UserList(Userindex).Char.Otra_Aura = 0
248             UserList(Userindex).Invent.Object(slot).Equipped = 0
250             UserList(Userindex).Invent.MagicoObjIndex = 0
252             UserList(Userindex).Invent.MagicoSlot = 0
        
254         Case eOBJType.otNUDILLOS
    
                'falta mandar animacion
            
256             UserList(Userindex).Invent.Object(slot).Equipped = 0
258             UserList(Userindex).Invent.NudilloObjIndex = 0
260             UserList(Userindex).Invent.NudilloSlot = 0
        
262             UserList(Userindex).Char.Arma_Aura = ""
264             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 1))
        
266             UserList(Userindex).Char.WeaponAnim = NingunArma
268             Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
        
270         Case eOBJType.otArmadura
272             UserList(Userindex).Invent.Object(slot).Equipped = 0
274             UserList(Userindex).Invent.ArmourEqpObjIndex = 0
276             UserList(Userindex).Invent.ArmourEqpSlot = 0
        
278             If UserList(Userindex).flags.Navegando = 0 Then
280                 If UserList(Userindex).flags.Montado = 0 Then
282                     Call DarCuerpoDesnudo(Userindex)
284                     Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                    End If
                End If
        
286             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 2))
        
294             UserList(Userindex).Char.Body_Aura = 0

                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If
    
296         Case eOBJType.otCASCO
298             UserList(Userindex).Invent.Object(slot).Equipped = 0
300             UserList(Userindex).Invent.CascoEqpObjIndex = 0
302             UserList(Userindex).Invent.CascoEqpSlot = 0
304             UserList(Userindex).Char.Head_Aura = 0
306             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 4))

308             UserList(Userindex).Char.CascoAnim = NingunCasco
310             Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
    
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If
    
318         Case eOBJType.otESCUDO
320             UserList(Userindex).Invent.Object(slot).Equipped = 0
322             UserList(Userindex).Invent.EscudoEqpObjIndex = 0
324             UserList(Userindex).Invent.EscudoEqpSlot = 0
326             UserList(Userindex).Char.Escudo_Aura = 0
328             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 3))
        
330             UserList(Userindex).Char.ShieldAnim = NingunEscudo

332             If UserList(Userindex).flags.Montado = 0 Then
334                 Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, UserList(Userindex).Char.Head, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                End If
                
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If
                
            Case eOBJType.otAnillos
                UserList(Userindex).Invent.Object(slot).Equipped = 0
                UserList(Userindex).Invent.AnilloEqpObjIndex = 0
                UserList(Userindex).Invent.AnilloEqpSlot = 0
                UserList(Userindex).Char.Anillo_Aura = 0
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(UserList(Userindex).Char.CharIndex, 0, True, 6))

                If obj.MagicDamageBonus > 0 Then
                    Call WriteUpdateDM(Userindex)
                End If
                
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If
        
        End Select

344     Call UpdateUserInv(False, Userindex, slot)

        
        Exit Sub

Desequipar_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.Desequipar", Erl)
        Resume Next
        
End Sub

Function SexoPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean

    On Error GoTo ErrHandler

    If EsGM(Userindex) Then
        SexoPuedeUsarItem = True
        Exit Function

    End If

    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(Userindex).genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(Userindex).genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True

    End If

    Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")

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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.FaccionPuedeUsarItem", Erl)
        Resume Next
        
End Function

Sub EquiparInvItem(ByVal Userindex As Integer, ByVal slot As Byte)

    On Error GoTo ErrHandler

    Dim errordesc As String

    'Equipa un item del inventario
    Dim obj       As ObjData
    Dim ObjIndex  As Integer

    ObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
    obj = ObjData(ObjIndex)

    If obj.Newbie = 1 And Not EsNewbie(Userindex) And Not EsGM(Userindex) Then
        Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(Userindex).Stats.ELV < obj.MinELV And Not EsGM(Userindex) Then
        Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If obj.SkillIndex > 0 Then
    
        If UserList(Userindex).Stats.UserSkills(obj.SkillIndex) < obj.SkillRequerido And Not EsGM(Userindex) Then
            Call WriteConsoleMsg(Userindex, "Necesitas " & obj.SkillRequerido & " puntos en " & SkillsNames(obj.SkillIndex) & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

    End If
    
    With UserList(Userindex)
    
        Select Case obj.OBJType

            Case eOBJType.otWeapon
                
                errordesc = "Arma"

                If Not ClasePuedeUsarItem(Userindex, ObjIndex, slot) And FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                    Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, slot)
                        
                    'Animacion por defecto
                    .Char.WeaponAnim = NingunArma

                    If .flags.Montado = 0 Then
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If

                    Exit Sub

                End If
            
                'Quitamos el elemento anterior
                If .Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
                End If
            
                If .Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.HerramientaEqpSlot)
                End If
            
                If .Invent.NudilloObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.NudilloSlot)
                End If
            
                .Invent.Object(slot).Equipped = 1
                .Invent.WeaponEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.WeaponEqpSlot = slot
            
                If obj.proyectil = 1 Then 'Si es un arco, desequipa el escudo.
            
                    'If .Invent.EscudoEqpObjIndex = 404 Or .Invent.EscudoEqpObjIndex = 1007 Or .Invent.EscudoEqpObjIndex = 1358 Then
                    If .Invent.EscudoEqpObjIndex = 1700 Or _
                       .Invent.EscudoEqpObjIndex = 1730 Or _
                       .Invent.EscudoEqpObjIndex = 1724 Or _
                       .Invent.EscudoEqpObjIndex = 1717 Or _
                       .Invent.EscudoEqpObjIndex = 1699 Then
                
                    Else

                        If .Invent.EscudoEqpObjIndex > 0 Then
                            Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
                            Call WriteConsoleMsg(Userindex, "No podes tirar flechas si tenés un escudo equipado. Tu escudo fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    End If

                End If
            
                'Sonido
                If obj.SndAura = 0 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                End If
            
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Arma_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                End If
                
                If obj.MagicDamageBonus > 0 Then
                    Call WriteUpdateDM(Userindex)
                End If
                
                If .flags.Montado = 0 Then
                
                    If .flags.Navegando = 0 Then
                        .Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                End If
      
            Case eOBJType.otHerramientas
        
                If Not ClasePuedeUsarItem(Userindex, ObjIndex, slot) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, slot)
                    Exit Sub

                End If

                If obj.MinSkill <> 0 Then
                
                    If .Stats.UserSkills(obj.QueSkill) < obj.MinSkill Then
                        Call WriteConsoleMsg(Userindex, "Para podes usar " & obj.name & " necesitas al menos " & obj.MinSkill & " puntos en " & SkillsNames(obj.QueSkill) & ".", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If

                End If

                'Quitamos el elemento anterior
                If .Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.HerramientaEqpSlot)
                End If
             
                If .Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
                End If
             
                .Invent.Object(slot).Equipped = 1
                .Invent.HerramientaEqpObjIndex = ObjIndex
                .Invent.HerramientaEqpSlot = slot
             
                If .flags.Montado = 0 Then
                
                    If .flags.Navegando = 0 Then
                        .Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                End If
       
            Case eOBJType.otmagicos
            
                errordesc = "Magico"
    
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If .Invent.MagicoObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.MagicoSlot)
                End If
        
                .Invent.Object(slot).Equipped = 1
                .Invent.MagicoObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.MagicoSlot = slot
                
                ' Debug.Print "magico" & obj.EfectoMagico
                Select Case obj.EfectoMagico

                    Case 1 ' Regenera Stamina
                        .flags.RegeneracionSta = 1

                    Case 2 'Modif la fuerza, agilidad, carisma, etc
                        ' .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo)
                        .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
                        
                        .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento
                        
                        If .Stats.UserAtributos(obj.QueAtributo) > MAXATRIBUTOS Then
                            .Stats.UserAtributos(obj.QueAtributo) = MAXATRIBUTOS
                        End If
                
                        Call WriteFYA(Userindex)

                    Case 3 'Modifica los skills
            
                        .Stats.UserSkills(obj.QueSkill) = .Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento

                    Case 4
                        .flags.RegeneracionHP = 1

                    Case 5
                        .flags.RegeneracionMana = 1

                    Case 6
                        'Call WriteConsoleMsg(UserIndex, "Item, temporalmente deshabilitado.", FontTypeNames.FONTTYPE_INFO)
                        .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
                        .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento

                    Case 9
                        .flags.NoMagiaEfeceto = 1

                    Case 10
                        .flags.incinera = 1

                    Case 11
                        .flags.Paraliza = 1

                    Case 12
                        .flags.CarroMineria = 1
                
                    Case 14
                        '.flags.DañoMagico = obj.CuantoAumento
                
                    Case 15 'Pendiete del Sacrificio
                        .flags.PendienteDelSacrificio = 1

                    Case 16
                        .flags.NoPalabrasMagicas = 1

                    Case 17
                        .flags.NoDetectable = 1
                   
                    Case 18 ' Pendiente del Experto
                        .flags.PendienteDelExperto = 1

                    Case 19
                        .flags.Envenena = 1

                    Case 20 'Anillo ocultismo
                        .flags.AnilloOcultismo = 1
    
                End Select
            
                'Sonido
                If obj.SndAura <> 0 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                End If
            
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Otra_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
                End If
        
                'Call WriteUpdateExp(UserIndex)
                'Call CheckUserLevel(UserIndex)
            
            Case eOBJType.otNUDILLOS
    
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If Not ClasePuedeUsarItem(Userindex, ObjIndex, slot) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                 
                If .Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.WeaponEqpSlot)

                End If

                If .Invent.Object(slot).Equipped Then
                    Call Desequipar(Userindex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If .Invent.NudilloObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.NudilloSlot)

                End If
        
                .Invent.Object(slot).Equipped = 1
                .Invent.NudilloObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.NudilloSlot = slot
        
                'Falta enviar anim
                If .flags.Montado = 0 Then
                
                    If .flags.Navegando = 0 Then
                        .Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                End If
            
                If obj.SndAura = 0 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.Y))
                End If
                 
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Arma_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
                End If
    
            Case eOBJType.otFlechas

                If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Or Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(Userindex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If .Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.MunicionEqpSlot)
                End If
        
                .Invent.Object(slot).Equipped = 1
                .Invent.MunicionEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.MunicionEqpSlot = slot

            Case eOBJType.otArmadura
                
                If obj.Ropaje = 0 Then
                    Call WriteConsoleMsg(Userindex, "Hay un error con este objeto. Infórmale a un administrador.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Nos aseguramos que puede usarla
                If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Or _
                   Not SexoPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Or _
                   Not CheckRazaUsaRopa(Userindex, .Invent.Object(slot).ObjIndex) Or _
                   Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
                    
                    Call WriteConsoleMsg(Userindex, "Tu clase, género, raza o facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    
                    Call Desequipar(Userindex, slot)

                    If .flags.Navegando = 0 Then
                        
                        If .flags.Montado = 0 Then
                            Call DarCuerpoDesnudo(Userindex)
                            Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                        End If

                    End If

                    Exit Sub

                End If

                'Quita el anterior
                If .Invent.ArmourEqpObjIndex > 0 Then
                    errordesc = "Armadura 2"
                    Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
                    errordesc = "Armadura 3"

                End If
  
                'Lo equipa
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Body_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))

                End If
            
                .Invent.Object(slot).Equipped = 1
                .Invent.ArmourEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.ArmourEqpSlot = slot
                            
                If .flags.Montado = 0 Then
                
                    If .flags.Navegando = 0 Then
                        
                        .Char.Body = obj.Ropaje
                
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        .flags.Desnudo = 0
            
                    End If

                End If
                
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If
    
            Case eOBJType.otCASCO
                
                If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
                    Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                    
                End If
                
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    Call Desequipar(Userindex, slot)
                
                    .Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    Exit Sub

                End If
    
                'Quita el anterior
                If .Invent.CascoEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.CascoEqpSlot)
                End If
            
                errordesc = "Casco"

                'Lo equipa
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Head_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
                End If
            
                .Invent.Object(slot).Equipped = 1
                .Invent.CascoEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.CascoEqpSlot = slot
            
                If .flags.Navegando = 0 Then
                    .Char.CascoAnim = obj.CascoAnim
                    Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End If
                
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If

            Case eOBJType.otESCUDO

                If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
                    Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    Call Desequipar(Userindex, slot)
                 
                    .Char.ShieldAnim = NingunEscudo

                    If .flags.Montado = 0 Then
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

                    End If

                    Exit Sub

                End If
     
                'Quita el anterior
                If .Invent.EscudoEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
                End If
     
                'Lo equipa
             
                If .Invent.Object(slot).ObjIndex = 1700 Or _
                   .Invent.Object(slot).ObjIndex = 1730 Or _
                   .Invent.Object(slot).ObjIndex = 1724 Or _
                   .Invent.Object(slot).ObjIndex = 1717 Or _
                   .Invent.Object(slot).ObjIndex = 1699 Then
             
                Else

                    If .Invent.WeaponEqpObjIndex > 0 Then
                        If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                            Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
                            Call WriteConsoleMsg(Userindex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    End If

                End If
            
                errordesc = "Escudo"
             
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Escudo_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
                End If

                .Invent.Object(slot).Equipped = 1
                .Invent.EscudoEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.EscudoEqpSlot = slot
                 
                If .flags.Navegando = 0 Then
                    If .flags.Montado = 0 Then
                        .Char.ShieldAnim = obj.ShieldAnim
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                End If
                
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If
                
            Case eOBJType.otAnillos

                If Not ClasePuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex, slot) Then
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                If Not FaccionPuedeUsarItem(Userindex, .Invent.Object(slot).ObjIndex) Then
                    Call WriteConsoleMsg(Userindex, "Tu facción no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                'Si esta equipado lo quita
                If .Invent.Object(slot).Equipped Then
                    Call Desequipar(Userindex, slot)
                    Exit Sub
                End If
     
                'Quita el anterior
                If .Invent.AnilloEqpSlot > 0 Then
                    Call Desequipar(Userindex, .Invent.AnilloEqpSlot)
                End If
                
                .Invent.Object(slot).Equipped = 1
                .Invent.AnilloEqpObjIndex = .Invent.Object(slot).ObjIndex
                .Invent.AnilloEqpSlot = slot
                
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Anillo_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Anillo_Aura, False, 6))
                End If

                If obj.MagicDamageBonus > 0 Then
                    Call WriteUpdateDM(Userindex)
                End If
                
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(Userindex)
                End If

        End Select
    
    End With

    'Actualiza
    Call UpdateUserInv(False, Userindex, slot)

    Exit Sub
    
ErrHandler:
    Debug.Print errordesc
    Call LogError("EquiparInvItem Slot:" & slot & " - Error: " & Err.Number & " - Error Description : " & Err.description & "- " & errordesc)

End Sub

Public Function CheckRazaUsaRopa(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean

    On Error GoTo ErrHandler

    If EsGM(Userindex) Then
        CheckRazaUsaRopa = True
        Exit Function

    End If

    Select Case UserList(Userindex).raza

        Case eRaza.Humano

            If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 And ObjData(ItemIndex).RazaDrow = 0 Then
                If ObjData(ItemIndex).Ropaje > 0 Then
                    CheckRazaUsaRopa = True
                    Exit Function

                End If

            End If

        Case eRaza.Elfo

            If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 And ObjData(ItemIndex).RazaDrow = 0 Then
                CheckRazaUsaRopa = True
                Exit Function

            End If
    
        Case eRaza.Orco

            If ObjData(ItemIndex).RazaEnana = 0 Then
                CheckRazaUsaRopa = True
                Exit Function

            End If
    
        Case eRaza.Drow

            If ObjData(ItemIndex).RazaEnana = 0 And ObjData(ItemIndex).RazaOrca = 0 Then
                CheckRazaUsaRopa = True
                Exit Function

            End If
    
        Case eRaza.Gnomo

            If ObjData(ItemIndex).RazaEnana > 0 Then
                CheckRazaUsaRopa = True
                Exit Function

            End If
        
        Case eRaza.Enano

            If ObjData(ItemIndex).RazaEnana > 0 Then
                CheckRazaUsaRopa = True
                Exit Function

            End If
    
    End Select

    CheckRazaUsaRopa = False

    Exit Function
ErrHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Public Function CheckRazaTipo(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean

    On Error GoTo ErrHandler

    If EsGM(Userindex) Then

        CheckRazaTipo = True
        Exit Function

    End If

    Select Case ObjData(ItemIndex).RazaTipo

        Case 0
            CheckRazaTipo = True

        Case 1

            If UserList(Userindex).raza = eRaza.Elfo Then
                CheckRazaTipo = True
                Exit Function

            End If
        
            If UserList(Userindex).raza = eRaza.Drow Then
                CheckRazaTipo = True
                Exit Function

            End If
        
            If UserList(Userindex).raza = eRaza.Humano Then
                CheckRazaTipo = True
                Exit Function

            End If

        Case 2

            If UserList(Userindex).raza = eRaza.Gnomo Then CheckRazaTipo = True
            If UserList(Userindex).raza = eRaza.Enano Then CheckRazaTipo = True
            Exit Function

        Case 3

            If UserList(Userindex).raza = eRaza.Orco Then CheckRazaTipo = True
            Exit Function
    
    End Select

    Exit Function
ErrHandler:
    Call LogError("Error CheckRazaTipo ItemIndex:" & ItemIndex)

End Function

Public Function CheckClaseTipo(ByVal Userindex As Integer, ItemIndex As Integer) As Boolean

    On Error GoTo ErrHandler

    If EsGM(Userindex) Then

        CheckClaseTipo = True
        Exit Function

    End If

    Select Case ObjData(ItemIndex).ClaseTipo

        Case 0
            CheckClaseTipo = True
            Exit Function

        Case 2

            If UserList(Userindex).clase = eClass.Mage Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Druid Then CheckClaseTipo = True
            Exit Function

        Case 1

            If UserList(Userindex).clase = eClass.Warrior Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Assasin Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Bard Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Cleric Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Paladin Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Trabajador Then CheckClaseTipo = True
            If UserList(Userindex).clase = eClass.Hunter Then CheckClaseTipo = True
            Exit Function

    End Select

    Exit Function
ErrHandler:
    Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)

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

    If UserList(Userindex).Invent.Object(slot).Amount = 0 Then Exit Sub

    obj = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex)

    If obj.OBJType = eOBJType.otWeapon Then
        If obj.proyectil = 1 Then

            'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
            If Not IntervaloPermiteUsar(Userindex, False) Then Exit Sub
        Else

            'dagas
            If Not IntervaloPermiteUsar(Userindex) Then Exit Sub

        End If

    Else

        If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
        If Not IntervaloPermiteGolpeUsar(Userindex, False) Then Exit Sub

    End If

    If UserList(Userindex).flags.Meditando Then
        UserList(Userindex).flags.Meditando = False
        UserList(Userindex).Char.FX = 0
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(UserList(Userindex).Char.CharIndex, 0))
    End If

    If obj.Newbie = 1 And Not EsNewbie(Userindex) And Not EsGM(Userindex) Then
        Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(Userindex).Stats.ELV < obj.MinELV Then
        Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    ObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
    UserList(Userindex).flags.TargetObjInvIndex = ObjIndex
    UserList(Userindex).flags.TargetObjInvSlot = slot

    Select Case obj.OBJType

        Case eOBJType.otUseOnce

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Usa el item
            UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam + obj.MinHam

            If UserList(Userindex).Stats.MinHam > UserList(Userindex).Stats.MaxHam Then UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam
            UserList(Userindex).flags.Hambre = 0
            Call WriteUpdateHungerAndThirst(Userindex)
            'Sonido
        
            If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

            End If
        
            'Quitamos del inv el item
            Call QuitarUserInvItem(Userindex, slot, 1)
        
            Call UpdateUserInv(False, Userindex, slot)

        Case eOBJType.otGuita

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + UserList(Userindex).Invent.Object(slot).Amount
            UserList(Userindex).Invent.Object(slot).Amount = 0
            UserList(Userindex).Invent.Object(slot).ObjIndex = 0
            UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
        
            Call UpdateUserInv(False, Userindex, slot)
            Call WriteUpdateGold(Userindex)
        
        Case eOBJType.otWeapon

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If Not UserList(Userindex).Stats.MinSta > 0 Then
                Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If ObjData(ObjIndex).proyectil = 1 Then
                'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
                Call WriteWorkRequestTarget(Userindex, Proyectiles)
            Else

                If UserList(Userindex).flags.TargetObj = Leña Then
                    If UserList(Userindex).Invent.Object(slot).ObjIndex = DAGA Then
                        Call TratarDeHacerFogata(UserList(Userindex).flags.TargetObjMap, UserList(Userindex).flags.TargetObjX, UserList(Userindex).flags.TargetObjY, Userindex)

                    End If

                End If

            End If
        
            'REVISAR LADDER
            'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
            If UserList(Userindex).Invent.Object(slot).Equipped = 0 Then
                'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                'Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
        Case eOBJType.otHerramientas

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If Not UserList(Userindex).Stats.MinSta > 0 Then
                Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
            If UserList(Userindex).Invent.Object(slot).Equipped = 0 Then
                'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(Userindex, "376", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            Select Case obj.Subtipo
                
                Case 1, 2  ' Herramientas del Pescador - Caña y Red
                    Call WriteWorkRequestTarget(Userindex, eSkill.Pescar)
                
                Case 3     ' Herramientas de Alquimia - Tijeras
                    Call WriteWorkRequestTarget(Userindex, eSkill.Alquimia)
                
                Case 4     ' Herramientas de Alquimia - Olla
                    Call EnivarObjConstruiblesAlquimia(Userindex)
                    Call WriteShowAlquimiaForm(Userindex)
                
                Case 5     ' Herramientas de Carpinteria - Serrucho
                    Call EnivarObjConstruibles(Userindex)
                    Call WriteShowCarpenterForm(Userindex)
                
                Case 6     ' Herramientas de Tala - Hacha
                    Call WriteWorkRequestTarget(Userindex, eSkill.Talar)

                Case 7     ' Herramientas de Herrero - Martillo
                    Call WriteConsoleMsg(Userindex, "Debes hacer click derecho sobre el yunque.", FontTypeNames.FONTTYPE_INFOIAO)

                Case 8     ' Herramientas de Mineria - Piquete
                    Call WriteWorkRequestTarget(Userindex, eSkill.Mineria)
                
                Case 9     ' Herramientas de Sastreria - Costurero
                    Call EnivarObjConstruiblesSastre(Userindex)
                    Call WriteShowSastreForm(Userindex)

            End Select
    
        Case eOBJType.otPociones

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            UserList(Userindex).flags.TomoPocion = True
            UserList(Userindex).flags.TipoPocion = obj.TipoPocion
                
            Dim CabezaFinal  As Integer

            Dim CabezaActual As Integer

            Select Case UserList(Userindex).flags.TipoPocion
        
                Case 1 'Modif la agilidad
                    UserList(Userindex).flags.DuracionEfecto = obj.DuracionEfecto
        
                    'Usa el item
                    UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                    
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(Userindex).Stats.UserAtributosBackUP(Agilidad) Then UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(Userindex).Stats.UserAtributosBackUP(Agilidad)
                
                    Call WriteFYA(Userindex)
                
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(Userindex, slot, 1)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If
        
                Case 2 'Modif la fuerza
                    UserList(Userindex).flags.DuracionEfecto = obj.DuracionEfecto
        
                    'Usa el item
                    UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                
                    If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(Userindex).Stats.UserAtributosBackUP(Fuerza) Then UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(Userindex).Stats.UserAtributosBackUP(Fuerza)
                
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(Userindex, slot, 1)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If

                    Call WriteFYA(Userindex)

                Case 3 'Pocion roja, restaura HP
                
                    'Usa el item
                    UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)

                    If UserList(Userindex).Stats.MinHp > UserList(Userindex).Stats.MaxHp Then UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
                
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(Userindex, slot, 1)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If
            
                Case 4 'Pocion azul, restaura MANA
            
                    Dim porcentajeRec As Byte
                    porcentajeRec = obj.Porcentaje
                
                    'Usa el item
                    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN + Porcentaje(UserList(Userindex).Stats.MaxMAN, porcentajeRec)

                    If UserList(Userindex).Stats.MinMAN > UserList(Userindex).Stats.MaxMAN Then UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
                
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(Userindex, slot, 1)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If
                
                Case 5 ' Pocion violeta

                    If UserList(Userindex).flags.Envenenado > 0 Then
                        UserList(Userindex).flags.Envenenado = 0
                        Call WriteConsoleMsg(Userindex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(Userindex, slot, 1)

                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

                    Else
                        Call WriteConsoleMsg(Userindex, "¡No te encuentras envenenado!", FontTypeNames.FONTTYPE_INFO)

                    End If
                
                Case 6  ' Remueve Parálisis

                    If UserList(Userindex).flags.Paralizado = 1 Or UserList(Userindex).flags.Inmovilizado = 1 Then
                        If UserList(Userindex).flags.Paralizado = 1 Then
                            UserList(Userindex).flags.Paralizado = 0
                            Call WriteParalizeOK(Userindex)

                        End If
                        
                        If UserList(Userindex).flags.Inmovilizado = 1 Then
                            UserList(Userindex).Counters.Inmovilizado = 0
                            UserList(Userindex).flags.Inmovilizado = 0
                            Call WriteInmovilizaOK(Userindex)

                        End If
                        
                        
                        
                        Call QuitarUserInvItem(Userindex, slot, 1)

                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(255, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

                        Call WriteConsoleMsg(Userindex, "Te has removido la paralizis.", FontTypeNames.FONTTYPE_INFOIAO)
                    Else
                        Call WriteConsoleMsg(Userindex, "No estas paralizado.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                
                Case 7  ' Pocion Naranja
                    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)

                    If UserList(Userindex).Stats.MinSta > UserList(Userindex).Stats.MaxSta Then UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
                    
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(Userindex, slot, 1)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If

                Case 8  ' Pocion cambio cara

                    Select Case UserList(Userindex).genero

                        Case eGenero.Hombre

                            Select Case UserList(Userindex).raza

                                Case eRaza.Humano
                                    CabezaFinal = RandomNumber(1, 40)

                                Case eRaza.Elfo
                                    CabezaFinal = RandomNumber(101, 132)

                                Case eRaza.Drow
                                    CabezaFinal = RandomNumber(201, 229)

                                Case eRaza.Enano
                                    CabezaFinal = RandomNumber(301, 329)

                                Case eRaza.Gnomo
                                    CabezaFinal = RandomNumber(401, 429)

                                Case eRaza.Orco
                                    CabezaFinal = RandomNumber(501, 529)

                            End Select

                        Case eGenero.Mujer

                            Select Case UserList(Userindex).raza

                                Case eRaza.Humano
                                    CabezaFinal = RandomNumber(50, 80)

                                Case eRaza.Elfo
                                    CabezaFinal = RandomNumber(150, 179)

                                Case eRaza.Drow
                                    CabezaFinal = RandomNumber(250, 279)

                                Case eRaza.Gnomo
                                    CabezaFinal = RandomNumber(350, 379)

                                Case eRaza.Enano
                                    CabezaFinal = RandomNumber(450, 479)

                                Case eRaza.Orco
                                    CabezaFinal = RandomNumber(550, 579)

                            End Select

                    End Select
            
                    UserList(Userindex).Char.Head = CabezaFinal
                    UserList(Userindex).OrigChar.Head = CabezaFinal
                    Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, CabezaFinal, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                    'Quitamos del inv el item
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 102, 0))

                    If CabezaActual <> CabezaFinal Then
                        Call QuitarUserInvItem(Userindex, slot, 1)
                    Else
                        Call WriteConsoleMsg(Userindex, "¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    
                Case 9  ' Pocion sexo
    
                    Select Case UserList(Userindex).genero

                        Case eGenero.Hombre
                            UserList(Userindex).genero = eGenero.Mujer
                    
                        Case eGenero.Mujer
                            UserList(Userindex).genero = eGenero.Hombre
                    
                    End Select
            
                    Select Case UserList(Userindex).genero

                        Case eGenero.Hombre

                            Select Case UserList(Userindex).raza

                                Case eRaza.Humano
                                    CabezaFinal = RandomNumber(1, 40)

                                Case eRaza.Elfo
                                    CabezaFinal = RandomNumber(101, 132)

                                Case eRaza.Drow
                                    CabezaFinal = RandomNumber(201, 229)

                                Case eRaza.Enano
                                    CabezaFinal = RandomNumber(301, 329)

                                Case eRaza.Gnomo
                                    CabezaFinal = RandomNumber(401, 429)

                                Case eRaza.Orco
                                    CabezaFinal = RandomNumber(501, 529)

                            End Select

                        Case eGenero.Mujer

                            Select Case UserList(Userindex).raza

                                Case eRaza.Humano
                                    CabezaFinal = RandomNumber(50, 80)

                                Case eRaza.Elfo
                                    CabezaFinal = RandomNumber(150, 179)

                                Case eRaza.Drow
                                    CabezaFinal = RandomNumber(250, 279)

                                Case eRaza.Gnomo
                                    CabezaFinal = RandomNumber(350, 379)

                                Case eRaza.Enano
                                    CabezaFinal = RandomNumber(450, 479)

                                Case eRaza.Orco
                                    CabezaFinal = RandomNumber(550, 579)

                            End Select

                    End Select
            
                    UserList(Userindex).Char.Head = CabezaFinal
                    UserList(Userindex).OrigChar.Head = CabezaFinal
                    Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, CabezaFinal, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)
                    'Quitamos del inv el item
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 102, 0))
                    Call QuitarUserInvItem(Userindex, slot, 1)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If
                
                Case 10  ' Invisibilidad
            
                    If UserList(Userindex).flags.invisible = 0 Then
                        UserList(Userindex).flags.invisible = 1
                        UserList(Userindex).Counters.Invisibilidad = obj.DuracionEfecto
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, True))
                        Call WriteContadores(Userindex)
                        Call QuitarUserInvItem(Userindex, slot, 1)

                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("123", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

                        Call WriteConsoleMsg(Userindex, "Te has escondido entre las sombras...", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        
                    Else
                        Call WriteConsoleMsg(Userindex, "Ya estas invisible.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        Exit Sub

                    End If
                    
                Case 11  ' Experiencia

                    Dim HR   As Integer

                    Dim MS   As Integer

                    Dim SS   As Integer

                    Dim secs As Integer

                    If UserList(Userindex).flags.ScrollExp = 1 Then
                        UserList(Userindex).flags.ScrollExp = obj.CuantoAumento
                        UserList(Userindex).Counters.ScrollExperiencia = obj.DuracionEfecto
                        Call QuitarUserInvItem(Userindex, slot, 1)
                        
                        secs = obj.DuracionEfecto
                        HR = secs \ 3600
                        MS = (secs Mod 3600) \ 60
                        SS = (secs Mod 3600) Mod 60

                        If SS > 9 Then
                            Call WriteConsoleMsg(Userindex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                        Else
                            Call WriteConsoleMsg(Userindex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)

                        End If

                    Else
                        Call WriteConsoleMsg(Userindex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                        Exit Sub

                    End If

                    Call WriteContadores(Userindex)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If

                Case 12  ' Oro
            
                    If UserList(Userindex).flags.ScrollOro = 1 Then
                        UserList(Userindex).flags.ScrollOro = obj.CuantoAumento
                        UserList(Userindex).Counters.ScrollOro = obj.DuracionEfecto
                        Call QuitarUserInvItem(Userindex, slot, 1)
                        secs = obj.DuracionEfecto
                        HR = secs \ 3600
                        MS = (secs Mod 3600) \ 60
                        SS = (secs Mod 3600) Mod 60

                        If SS > 9 Then
                            Call WriteConsoleMsg(Userindex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                        Else
                            Call WriteConsoleMsg(Userindex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)

                        End If
                        
                    Else
                        Call WriteConsoleMsg(Userindex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                        Exit Sub

                    End If

                    Call WriteContadores(Userindex)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If

                Case 13
                
                    Call QuitarUserInvItem(Userindex, slot, 1)
                    UserList(Userindex).flags.Envenenado = 0
                    UserList(Userindex).flags.Incinerado = 0
                    
                    If UserList(Userindex).flags.Inmovilizado = 1 Then
                        UserList(Userindex).Counters.Inmovilizado = 0
                        UserList(Userindex).flags.Inmovilizado = 0
                        Call WriteInmovilizaOK(Userindex)
                        

                    End If
                    
                    If UserList(Userindex).flags.Paralizado = 1 Then
                        UserList(Userindex).flags.Paralizado = 0
                        Call WriteParalizeOK(Userindex)
                        

                    End If
                    
                    If UserList(Userindex).flags.Ceguera = 1 Then
                        UserList(Userindex).flags.Ceguera = 0
                        Call WriteBlindNoMore(Userindex)
                        

                    End If
                    
                    If UserList(Userindex).flags.Maldicion = 1 Then
                        UserList(Userindex).flags.Maldicion = 0
                        UserList(Userindex).Counters.Maldicion = 0

                    End If
                    
                    UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta
                    UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
                    UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MaxMAN
                    UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
                    UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MaxHam
                    
                    UserList(Userindex).flags.Hambre = 0
                    UserList(Userindex).flags.Sed = 0
                    
                    Call WriteUpdateHungerAndThirst(Userindex)
                    Call WriteConsoleMsg(Userindex, "Donador> Te sentis sano y lleno.", FontTypeNames.FONTTYPE_WARNING)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If

                Case 14
                
                    If UserList(Userindex).flags.BattleModo = 1 Then
                        Call WriteConsoleMsg(Userindex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub

                    End If
                    
                    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = CARCEL Then
                        Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
                    Dim Map     As Integer

                    Dim X       As Byte

                    Dim Y       As Byte

                    Dim DeDonde As WorldPos

                    Call QuitarUserInvItem(Userindex, slot, 1)
            
                    Select Case UserList(Userindex).Hogar

                        Case eCiudad.cUllathorpe
                            DeDonde = Ullathorpe
                            
                        Case eCiudad.cNix
                            DeDonde = Nix
                
                        Case eCiudad.cBanderbill
                            DeDonde = Banderbill
                        
                        Case eCiudad.cLindos
                            DeDonde = Lindos
                            
                        Case eCiudad.cArghal
                            DeDonde = Arghal
                            
                        Case eCiudad.CHillidan
                            DeDonde = Hillidan
                            
                        Case Else
                            DeDonde = Ullathorpe

                    End Select
                    
                    Map = DeDonde.Map
                    X = DeDonde.X
                    Y = DeDonde.Y
                    
                    Call FindLegalPos(Userindex, Map, X, Y)
                    Call WarpUserChar(Userindex, Map, X, Y, True)
                    Call WriteConsoleMsg(Userindex, "Ya estas a salvo...", FontTypeNames.FONTTYPE_WARNING)

                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If

                Case 15  ' Aliento de sirena
                        
                    If UserList(Userindex).Counters.Oxigeno >= 3540 Then
                        
                        Call WriteConsoleMsg(Userindex, "No podes acumular más de 59 minutos de oxigeno.", FontTypeNames.FONTTYPE_INFOIAO)
                        secs = UserList(Userindex).Counters.Oxigeno
                        HR = secs \ 3600
                        MS = (secs Mod 3600) \ 60
                        SS = (secs Mod 3600) Mod 60

                        If SS > 9 Then
                            Call WriteConsoleMsg(Userindex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                        Else
                            Call WriteConsoleMsg(Userindex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)

                        End If

                    Else
                            
                        UserList(Userindex).Counters.Oxigeno = UserList(Userindex).Counters.Oxigeno + obj.DuracionEfecto
                        Call QuitarUserInvItem(Userindex, slot, 1)
                            
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
                            
                        UserList(Userindex).flags.Ahogandose = 0
                        Call WriteOxigeno(Userindex)
                            
                        Call WriteContadores(Userindex)

                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If

                    End If

                Case 16 ' Divorcio

                    If UserList(Userindex).flags.Casado = 1 Then

                        Dim tUser As Integer

                        'UserList(UserIndex).flags.Pareja
                        tUser = NameIndex(UserList(Userindex).flags.Pareja)
                        Call QuitarUserInvItem(Userindex, slot, 1)
                        
                        If tUser <= 0 Then

                            Dim FileUser As String

                            FileUser = CharPath & UCase$(UserList(Userindex).flags.Pareja) & ".chr"
                            'Call WriteVar(FileUser, "FLAGS", "CASADO", 0)
                            'Call WriteVar(FileUser, "FLAGS", "PAREJA", "")
                            UserList(Userindex).flags.Casado = 0
                            UserList(Userindex).flags.Pareja = ""
                            Call WriteConsoleMsg(Userindex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
                            UserList(Userindex).MENSAJEINFORMACION = UserList(Userindex).name & " se ha divorciado de ti."

                        Else
                            UserList(tUser).flags.Casado = 0
                            UserList(tUser).flags.Pareja = ""
                            UserList(Userindex).flags.Casado = 0
                            UserList(Userindex).flags.Pareja = ""
                            Call WriteConsoleMsg(Userindex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteConsoleMsg(tUser, UserList(Userindex).name & " se ha divorciado de ti.", FontTypeNames.FONTTYPE_INFOIAO)
                            
                        End If

                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                        End If
                    
                    Else
                        Call WriteConsoleMsg(Userindex, "No estas casado.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                Case 17 'Cara legendaria

                    Select Case UserList(Userindex).genero

                        Case eGenero.Hombre

                            Select Case UserList(Userindex).raza

                                Case eRaza.Humano
                                    CabezaFinal = RandomNumber(684, 686)

                                Case eRaza.Elfo
                                    CabezaFinal = RandomNumber(690, 692)

                                Case eRaza.Drow
                                    CabezaFinal = RandomNumber(696, 698)

                                Case eRaza.Enano
                                    CabezaFinal = RandomNumber(702, 704)

                                Case eRaza.Gnomo
                                    CabezaFinal = RandomNumber(708, 710)

                                Case eRaza.Orco
                                    CabezaFinal = RandomNumber(714, 716)

                            End Select

                        Case eGenero.Mujer

                            Select Case UserList(Userindex).raza

                                Case eRaza.Humano
                                    CabezaFinal = RandomNumber(687, 689)

                                Case eRaza.Elfo
                                    CabezaFinal = RandomNumber(693, 695)

                                Case eRaza.Drow
                                    CabezaFinal = RandomNumber(699, 701)

                                Case eRaza.Gnomo
                                    CabezaFinal = RandomNumber(705, 707)

                                Case eRaza.Enano
                                    CabezaFinal = RandomNumber(711, 713)

                                Case eRaza.Orco
                                    CabezaFinal = RandomNumber(717, 719)

                            End Select

                    End Select

                    CabezaActual = UserList(Userindex).OrigChar.Head
                        
                    UserList(Userindex).Char.Head = CabezaFinal
                    UserList(Userindex).OrigChar.Head = CabezaFinal
                    Call ChangeUserChar(Userindex, UserList(Userindex).Char.Body, CabezaFinal, UserList(Userindex).Char.Heading, UserList(Userindex).Char.WeaponAnim, UserList(Userindex).Char.ShieldAnim, UserList(Userindex).Char.CascoAnim)

                    'Quitamos del inv el item
                    If CabezaActual <> CabezaFinal Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 102, 0))
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        Call QuitarUserInvItem(Userindex, slot, 1)
                    Else
                        Call WriteConsoleMsg(Userindex, "¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                Case 18  ' tan solo crea una particula por determinado tiempo

                    Dim Particula           As Integer

                    Dim Tiempo              As Long

                    Dim ParticulaPermanente As Byte

                    Dim sobrechar           As Byte

                    If obj.CreaParticula <> "" Then
                        Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
                        Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
                        ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
                        sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                            
                        If ParticulaPermanente = 1 Then
                            UserList(Userindex).Char.ParticulaFx = Particula
                            UserList(Userindex).Char.loops = Tiempo

                        End If
                            
                        If sobrechar = 1 Then
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, Particula, Tiempo))
                        Else
                            
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, Particula, Tiempo, False))

                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, Particula, Tiempo))
                        End If

                    End If
                        
                    If obj.CreaFX <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageFxPiso(obj.CreaFX, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                            
                        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, obj.CreaFX, 0))
                        ' PrepareMessageCreateFX
                    End If
                        
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

                    End If
                        
                    Call QuitarUserInvItem(Userindex, slot, 1)

                Case 19 ' Reseteo de skill

                    Dim S As Byte
                
                    If UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) >= 80 Then
                        Call WriteConsoleMsg(Userindex, "Has fundado un clan, no podes resetar tus skills. ", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
                    
                    For S = 1 To NUMSKILLS
                        UserList(Userindex).Stats.UserSkills(S) = 0
                    Next S
                    
                    Dim SkillLibres As Integer
                    
                    SkillLibres = 5
                    SkillLibres = SkillLibres + (5 * UserList(Userindex).Stats.ELV)
                     
                    UserList(Userindex).Stats.SkillPts = SkillLibres
                    Call WriteLevelUp(Userindex, UserList(Userindex).Stats.SkillPts)
                    
                    Call WriteConsoleMsg(Userindex, "Tus skills han sido reseteados.", FontTypeNames.FONTTYPE_INFOIAO)
                    Call QuitarUserInvItem(Userindex, slot, 1)

                Case 20
                
                    If UserList(Userindex).Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
                        UserList(Userindex).Stats.InventLevel = UserList(Userindex).Stats.InventLevel + 1
                        UserList(Userindex).CurrentInventorySlots = getMaxInventorySlots(Userindex)
                        Call WriteInventoryUnlockSlots(Userindex)
                        Call WriteConsoleMsg(Userindex, "Has aumentado el espacio de tu inventario!", FontTypeNames.FONTTYPE_INFO)
                        Call QuitarUserInvItem(Userindex, slot, 1)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ya has desbloqueado todos los casilleros disponibles.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
            End Select

            Call WriteUpdateUserStats(Userindex)
            Call UpdateUserInv(False, Userindex, slot)

        Case eOBJType.otBebidas

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + obj.MinSed

            If UserList(Userindex).Stats.MinAGU > UserList(Userindex).Stats.MaxAGU Then UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
            UserList(Userindex).flags.Sed = 0
            Call WriteUpdateHungerAndThirst(Userindex)
        
            'Quitamos del inv el item
            Call QuitarUserInvItem(Userindex, slot, 1)
        
            If obj.Snd1 <> 0 Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
            
            Else
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

            End If
        
            Call UpdateUserInv(False, Userindex, slot)
        
        Case eOBJType.OtCofre

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            'Quitamos del inv el item
            Call QuitarUserInvItem(Userindex, slot, 1)
            Call UpdateUserInv(False, Userindex, slot)
        
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg(UserList(Userindex).name & " ha abierto un " & obj.name & " y obtuvo...", FontTypeNames.FONTTYPE_New_DONADOR))
        
            If obj.Snd1 <> 0 Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))

            End If
        
            If obj.CreaFX <> 0 Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, obj.CreaFX, 0))

            End If
        
            Dim i As Byte

            If obj.Subtipo = 1 Then

                For i = 1 To obj.CantItem

                    If Not MeterItemEnInventario(Userindex, obj.Item(i)) Then Call TirarItemAlPiso(UserList(Userindex).Pos, obj.Item(i))
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
                Next i
        
            Else
        
                For i = 1 To obj.CantEntrega

                    Dim indexobj As Byte
                
                    indexobj = RandomNumber(1, obj.CantItem)
            
                    Dim Index As obj

                    Index.ObjIndex = obj.Item(indexobj).ObjIndex
                    Index.Amount = obj.Item(indexobj).Amount

                    If Not MeterItemEnInventario(Userindex, Index) Then Call TirarItemAlPiso(UserList(Userindex).Pos, Index)
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).name & " (" & Index.Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
                Next i

            End If
    
        Case eOBJType.otLlaves
            Call WriteConsoleMsg(Userindex, "Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.", FontTypeNames.FONTTYPE_INFO)
    
        Case eOBJType.otBotellaVacia

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If (MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
                Call WriteConsoleMsg(Userindex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(Userindex, slot, 1)

            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

            End If
        
            Call UpdateUserInv(False, Userindex, slot)
    
        Case eOBJType.otBotellaLlena

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU + obj.MinSed

            If UserList(Userindex).Stats.MinAGU > UserList(Userindex).Stats.MaxAGU Then UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MaxAGU
            UserList(Userindex).flags.Sed = 0
            Call WriteUpdateHungerAndThirst(Userindex)
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(Userindex).Invent.Object(slot).ObjIndex).IndexCerrada
            Call QuitarUserInvItem(Userindex, slot, 1)

            If Not MeterItemEnInventario(Userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(Userindex).Pos, MiObj)

            End If
        
            Call UpdateUserInv(False, Userindex, slot)
    
        Case eOBJType.otPergaminos

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Call LogError(UserList(UserIndex).Name & " intento aprender el hechizo " & ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex)
        
            If ClasePuedeUsarItem(Userindex, UserList(Userindex).Invent.Object(slot).ObjIndex, slot) Then

                'If UserList(UserIndex).Stats.MaxMAN > 0 Then
                If UserList(Userindex).flags.Hambre = 0 And UserList(Userindex).flags.Sed = 0 Then
                    Call AgregarHechizo(Userindex, slot)
                    Call UpdateUserInv(False, Userindex, slot)
                    ' Call LogError(UserList(UserIndex).Name & " lo aprendio.")
                Else
                    Call WriteConsoleMsg(Userindex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)

                End If

                ' Else
                '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_WARNING)
                'End If
            Else
             
                Call WriteConsoleMsg(Userindex, "Por mas que lo intentas, no podés comprender el manuescrito.", FontTypeNames.FONTTYPE_INFO)
   
            End If
        
        Case eOBJType.otMinerales

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            Call WriteWorkRequestTarget(Userindex, FundirMetal)
       
        Case eOBJType.otInstrumentos

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If obj.Real Then '¿Es el Cuerno Real?
                If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                    If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
                        Call WriteConsoleMsg(Userindex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call SendData(SendTarget.toMap, UserList(Userindex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    Exit Sub
                Else
                    Call WriteConsoleMsg(Userindex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            ElseIf obj.Caos Then '¿Es el Cuerno Legión?

                If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                    If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
                        Call WriteConsoleMsg(Userindex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Call SendData(SendTarget.toMap, UserList(Userindex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                    Exit Sub
                Else
                    Call WriteConsoleMsg(Userindex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If

            'Si llega aca es porque es o Laud o Tambor o Flauta
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
       
        Case eOBJType.otBarcos
            
            'Verifica si tiene el nivel requerido para navegar, siendo Trabajador o Pirata
            If UserList(Userindex).Stats.ELV < 20 And (UserList(Userindex).clase = eClass.Trabajador Or UserList(Userindex).clase = eClass.Pirat) Then
                Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            'Verifica si tiene el nivel requerido para navegar, sin ser Trabajador o Pirata
            ElseIf UserList(Userindex).Stats.ELV < 25 Then
                Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If obj.Subtipo = 0 Then
            If UserList(Userindex).flags.Navegando = 0 Then
                If ((LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X - 1, UserList(Userindex).Pos.Y, True, False) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1, True, False) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X + 1, UserList(Userindex).Pos.Y, True, False) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, True, False)) And UserList(Userindex).flags.Navegando = 0) Or UserList(Userindex).flags.Navegando = 1 Then
                    Call DoNavega(Userindex, obj, slot)
                Else
                    Call WriteConsoleMsg(Userindex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)

                End If

            Else 'Ladder 10-02-2010

                If UserList(Userindex).Invent.BarcoObjIndex <> UserList(Userindex).Invent.Object(slot).ObjIndex Then
                    Call DoReNavega(Userindex, obj, slot)
                Else

                    If ((LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X - 1, UserList(Userindex).Pos.Y, False, True) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y - 1, False, True) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X + 1, UserList(Userindex).Pos.Y, False, True) Or LegalPos(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y + 1, False, True)) And UserList(Userindex).flags.Navegando = 1) Or UserList(Userindex).flags.Navegando = 0 Then
                        Call DoNavega(Userindex, obj, slot)
                    Else
                        Call WriteConsoleMsg(Userindex, "¡Debes aproximarte a la costa para dejar la barca!", FontTypeNames.FONTTYPE_INFO)

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
        
        Case eOBJType.otMonturas
            'Verifica todo lo que requiere la montura
    
            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡Estas muerto! Los fantasmas no pueden montar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If UserList(Userindex).flags.Navegando = 1 Then
                Call WriteConsoleMsg(Userindex, "Debes dejar de navegar para poder montarté.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If MapInfo(UserList(Userindex).Pos.Map).zone = "DUNGEON" Then
                Call WriteConsoleMsg(Userindex, "No podes cabalgar dentro de un dungeon.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Call DoMontar(Userindex, obj, slot)

        Case eOBJType.OtDonador

            Select Case obj.Subtipo

                Case 1
            
                    If UserList(Userindex).Counters.Pena <> 0 Then
                        Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
                    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = CARCEL Then
                        Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
                    Call WarpUserChar(Userindex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                    Call WriteConsoleMsg(Userindex, "Has viajado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
                    Call QuitarUserInvItem(Userindex, slot, 1)
                    Call UpdateUserInv(False, Userindex, slot)
                
                Case 2

                    If DonadorCheck(UserList(Userindex).Cuenta) = 0 Then
                        Call DonadorTiempo(UserList(Userindex).Cuenta, CLng(obj.CuantoAumento))
                        Call WriteConsoleMsg(Userindex, "Donación> Se han agregado " & obj.CuantoAumento & " dias de donador a tu cuenta. Relogea tu personaje para empezar a disfrutar la experiencia.", FontTypeNames.FONTTYPE_WARNING)
                        Call QuitarUserInvItem(Userindex, slot, 1)
                        Call UpdateUserInv(False, Userindex, slot)
                    Else
                        Call DonadorTiempo(UserList(Userindex).Cuenta, CLng(obj.CuantoAumento))
                        Call WriteConsoleMsg(Userindex, "¡Se han añadido " & CLng(obj.CuantoAumento) & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
                        UserList(Userindex).donador.activo = 1
                        Call QuitarUserInvItem(Userindex, slot, 1)
                        Call UpdateUserInv(False, Userindex, slot)

                        'Call WriteConsoleMsg(UserIndex, "Donación> Debes esperar a que finalice el periodo existente para renovar tu suscripción.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If

                Case 3
                    Call AgregarCreditosDonador(UserList(Userindex).Cuenta, CLng(obj.CuantoAumento))
                    Call WriteConsoleMsg(Userindex, "Donación> Tu credito ahora es de " & CreditosDonadorCheck(UserList(Userindex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                    Call QuitarUserInvItem(Userindex, slot, 1)
                    Call UpdateUserInv(False, Userindex, slot)

            End Select
     
        Case eOBJType.otpasajes

            If UserList(Userindex).flags.Muerto = 1 Then
                Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If UserList(Userindex).flags.TargetNpcTipo <> Pirata Then
                Call WriteConsoleMsg(Userindex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If Distancia(Npclist(UserList(Userindex).flags.TargetNPC).Pos, UserList(Userindex).Pos) > 3 Then
                Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If UserList(Userindex).Pos.Map <> obj.DesdeMap Then
                Rem  Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aquí! Largate!", FontTypeNames.FONTTYPE_INFO)
                Call WriteChatOverHead(Userindex, "El pasaje no lo compraste aquí! Largate!", str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            End If
        
            If Not MapaValido(obj.HastaMap) Then
                Rem Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
                Call WriteChatOverHead(Userindex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            End If

            If obj.NecesitaNave > 0 Then
                If UserList(Userindex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                    Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteChatOverHead(Userindex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub

                End If

            End If
            
            Call WarpUserChar(Userindex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
            Call WriteConsoleMsg(Userindex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
            UserList(Userindex).Stats.MinAGU = 0
            UserList(Userindex).Stats.MinHam = 0
            UserList(Userindex).flags.Sed = 1
            UserList(Userindex).flags.Hambre = 1
            Call WriteUpdateHungerAndThirst(Userindex)
            Call QuitarUserInvItem(Userindex, slot, 1)
            Call UpdateUserInv(False, Userindex, slot)
        
        Case eOBJType.otRunas
    
            If UserList(Userindex).Counters.Pena <> 0 Then
                Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = CARCEL Then
                Call WriteConsoleMsg(Userindex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If UserList(Userindex).flags.BattleModo = 1 Then
                Call WriteConsoleMsg(Userindex, "No podes usarlo aquí.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
            If MapInfo(UserList(Userindex).Pos.Map).Seguro = 0 And UserList(Userindex).flags.Muerto = 0 Then
                Call WriteConsoleMsg(Userindex, "Solo podes usar tu runa en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If UserList(Userindex).Accion.AccionPendiente Then
                Exit Sub

            End If
        
            Select Case ObjData(ObjIndex).TipoRuna
        
                Case 1, 2

                    If UserList(Userindex).donador.activo = 0 Then ' Donador no espera tiempo
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 350, Accion_Barra.Runa))
                    Else
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 100, Accion_Barra.Runa))

                    End If

                    UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
                    UserList(Userindex).Accion.AccionPendiente = True
                    UserList(Userindex).Accion.TipoAccion = Accion_Barra.Runa
                    UserList(Userindex).Accion.RunaObj = ObjIndex
                    UserList(Userindex).Accion.ObjSlot = slot
            
                Case 3
        
                    Dim parejaindex As Integer

                    If Not UserList(Userindex).flags.BattleModo Then
                
                        'If UserList(UserIndex).donador.activo = 1 Then
                        If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
                            If UserList(Userindex).flags.Casado = 1 Then
                                parejaindex = NameIndex(UserList(Userindex).flags.Pareja)
                        
                                If parejaindex > 0 Then
                                    If UserList(parejaindex).flags.BattleModo = 0 Then
                                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 600, False))
                                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 600, Accion_Barra.GoToPareja))
                                        UserList(Userindex).Accion.AccionPendiente = True
                                        UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
                                        UserList(Userindex).Accion.TipoAccion = Accion_Barra.GoToPareja
                                    Else
                                        Call WriteConsoleMsg(Userindex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                                    End If
                                
                                Else
                                    Call WriteConsoleMsg(Userindex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                                End If

                            Else
                                Call WriteConsoleMsg(Userindex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        Else
                            Call WriteConsoleMsg(Userindex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If
                
                        ' Else
                        '  Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
                        ' End If
                    Else
                        Call WriteConsoleMsg(Userindex, "No podés usar esta opción en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
        
                    End If
    
            End Select
        
        Case eOBJType.otmapa
            Call WriteShowFrmMapa(Userindex)
        
    End Select

    Exit Sub

hErr:
    LogError "Error en useinvitem Usuario: " & UserList(Userindex).name & " item:" & obj.name & " index: " & UserList(Userindex).Invent.Object(slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal Userindex As Integer)
        
        On Error GoTo EnivarArmasConstruibles_Err
        

100     Call WriteBlacksmithWeapons(Userindex)

        
        Exit Sub

EnivarArmasConstruibles_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarArmasConstruibles", Erl)
        Resume Next
        
End Sub
 
Sub EnivarObjConstruibles(ByVal Userindex As Integer)
        
        On Error GoTo EnivarObjConstruibles_Err
        

100     Call WriteCarpenterObjects(Userindex)

        
        Exit Sub

EnivarObjConstruibles_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruibles", Erl)
        Resume Next
        
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal Userindex As Integer)
        
        On Error GoTo EnivarObjConstruiblesAlquimia_Err
        

100     Call WriteAlquimistaObjects(Userindex)

        
        Exit Sub

EnivarObjConstruiblesAlquimia_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)
        Resume Next
        
End Sub

Sub EnivarObjConstruiblesSastre(ByVal Userindex As Integer)
        
        On Error GoTo EnivarObjConstruiblesSastre_Err
        

100     Call WriteSastreObjects(Userindex)

        
        Exit Sub

EnivarObjConstruiblesSastre_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)
        Resume Next
        
End Sub

Sub EnivarArmadurasConstruibles(ByVal Userindex As Integer)
        
        On Error GoTo EnivarArmadurasConstruibles_Err
        

100     Call WriteBlacksmithArmors(Userindex)

        
        Exit Sub

EnivarArmadurasConstruibles_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.EnivarArmadurasConstruibles", Erl)
        Resume Next
        
End Sub

Sub TirarTodo(ByVal Userindex As Integer)

    On Error Resume Next

    If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 6 Then Exit Sub
    If UserList(Userindex).flags.BattleModo = 1 Then Exit Sub

    Call TirarTodosLosItems(Userindex)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemSeCae_Err
        

100     ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> eOBJType.otLlaves And ObjData(Index).OBJType <> eOBJType.otBarcos And ObjData(Index).OBJType <> eOBJType.otMonturas And ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And ObjData(Index).donador = 0 And Not ObjData(Index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.ItemSeCae", Erl)
        Resume Next
        
End Function

Public Function PirataCaeItem(ByVal Userindex As Integer, ByVal slot As Byte)

    With UserList(Userindex)
    
        If .clase = eClass.Pirat Then
            
            ' El pirata con galera no pierde los últimos 6 * (cada 10 niveles; max 1) slots
            If ObjData(.Invent.BarcoObjIndex).Ropaje = iGalera Then
            
                If slot > .CurrentInventorySlots - 6 * min(.Stats.ELV \ 10, 1) Then
                    Exit Function
                End If
            
            ' Con galeón no pierde los últimos 6 * (cada 10 niveles; max 3) slots
            ElseIf ObjData(.Invent.BarcoObjIndex).Ropaje = iGaleon Then
            
                If slot > .CurrentInventorySlots - 6 * min(.Stats.ELV \ 10, 3) Then
                    Exit Function
                End If
            
            End If
            
        End If
        
    End With
    
    PirataCaeItem = True

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
    
    With UserList(Userindex)
    
        For i = 1 To .CurrentInventorySlots
    
            ItemIndex = .Invent.Object(i).ObjIndex

            If ItemIndex > 0 Then

                If ItemSeCae(ItemIndex) And PirataCaeItem(Userindex, i) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                
                    If .flags.CarroMineria = 1 Then
                
                        If ItemIndex = ORO_MINA Or ItemIndex = PLATA_MINA Or ItemIndex = HIERRO_MINA Then
                       
                            MiObj.Amount = .Invent.Object(i).Amount * 0.3
                            MiObj.ObjIndex = ItemIndex
                        
                            Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                    
                            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                                Call DropObj(Userindex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                            End If

                        End If
                    
                    Else
                    
                        MiObj.Amount = .Invent.Object(i).Amount
                        MiObj.ObjIndex = ItemIndex
                        
                        Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                
                        If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                            Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        End If
                    
                    End If
                
                End If

            End If
    
        Next i
    
    End With
 
    Exit Sub

TirarTodosLosItems_Err:
    Call RegistrarError(Err.Number, Err.description, "InvUsuario.TirarTodosLosItems", Erl)

    Resume Next
        
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo ItemNewbie_Err
        

100     ItemNewbie = ObjData(ItemIndex).Newbie = 1

        
        Exit Function

ItemNewbie_Err:
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.ItemNewbie", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "InvUsuario.TirarTodosLosItemsNoNewbies", Erl)
        Resume Next
        
End Sub
