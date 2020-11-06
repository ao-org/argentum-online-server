Attribute VB_Name = "InvUsuario"
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la funci�n se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer


For i = 1 To UserList(UserIndex).CurrentInventorySlots
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And ObjData(ObjIndex).OBJType <> eOBJType.OtDonador And ObjData(ObjIndex).OBJType <> eOBJType.otRunas) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional slot As Byte) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

If slot <> 0 Then
If UserList(UserIndex).Invent.Object(slot).Equipped Then
ClasePuedeUsarItem = True
Exit Function
End If
End If

If EsGM(UserIndex) Then
    ClasePuedeUsarItem = True
    Exit Function
End If


'Admins can use ANYTHING!
'If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
    'If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
        Dim i As Integer
        For i = 1 To 9
            If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
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

Sub QuitarNewbieObj(ByVal UserIndex As Integer)

    Dim j As Integer
    For j = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
            If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then
                Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                Call UpdateUserInv(False, UserIndex, j)
            End If
        
        End If
    Next j
    
    'Si el usuario dej� de ser Newbie, y estaba en el Newbie Dungeon
    'es transportado a su hogar de origen ;)
    If UCase$(MapInfo(UserList(UserIndex).Pos.Map).restrict_mode) = "NEWBIE" Then
        
        Dim DeDonde As WorldPos
        
        Select Case UserList(UserIndex).Hogar
            Case eCiudad.cUllathorpe
                DeDonde = Ullathorpe
                
            Case eCiudad.cNix
                DeDonde = Nix
    
            Case eCiudad.cBanderbill
                DeDonde = Banderbill
            
            Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                DeDonde = Lindos
                
            Case eCiudad.cArghal 'Vamos a tener que ir por todo el desierto... uff!
                DeDonde = Arghal
                
            Case eCiudad.CHillidan
                DeDonde = Hillidan
                
            Case Else
                DeDonde = Ullathorpe
        End Select
        
        Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.x, DeDonde.Y, True)
    
    End If

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To UserList(UserIndex).CurrentInventorySlots
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).Amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
UserList(UserIndex).Invent.HerramientaEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0



UserList(UserIndex).Invent.NudilloObjIndex = 0
UserList(UserIndex).Invent.NudilloSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

UserList(UserIndex).Invent.MonturaObjIndex = 0
UserList(UserIndex).Invent.MonturaSlot = 0


UserList(UserIndex).Invent.MagicoObjIndex = 0
UserList(UserIndex).Invent.MagicoSlot = 0



End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo Errhandler

'If Cantidad > 100000 Then Exit Sub
If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub
'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As obj
        Dim Logs As Long
        'info debug
        Dim loops As Integer
        
        Logs = Cantidad

        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.Amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.Amount
            Else
                MiObj.Amount = Cantidad
                Cantidad = Cantidad - MiObj.Amount
            End If

            MiObj.ObjIndex = iORO
            

            Dim AuxPos As WorldPos
            

                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                
                
                If AuxPos.x <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.Amount
                End If
        
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
            
        Loop
        
        
        If EsGM(UserIndex) Then
                If MiObj.ObjIndex = iORO Then
                    Call LogGM(UserList(UserIndex).name, "Tiro: " & Logs & " monedas de oro.")
                Else
                    Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & Logs & " Objeto:" & ObjData(MiObj.ObjIndex).name)
                End If
            End If
        
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
    If slot < 1 Or slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(slot)
        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, slot)
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '�Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
    End With
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte
'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, slot, UserList(UserIndex).Invent.Object(slot))
    Else
        Call ChangeUserInv(UserIndex, slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)



    Dim obj As obj
    If num > 0 Then
      
      If num > UserList(UserIndex).Invent.Object(slot).Amount Then num = UserList(UserIndex).Invent.Object(slot).Amount
        obj.ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
        obj.Amount = num
        If ObjData(obj.ObjIndex).Destruye = 0 Then
            'Check objeto en el suelo
            If MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex = 0 Then
                  
                  
                  If num + MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
                      num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount
                  End If
                  
                  
                  
                  Call MakeObj(obj, Map, x, Y)
                  Call QuitarUserInvItem(UserIndex, slot, num)
                  Call UpdateUserInv(False, UserIndex, slot)
                  
            
                  
                  If Not UserList(UserIndex).flags.Privilegios And PlayerType.user Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)
                  
                  'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
                  'Es un Objeto que tenemos que loguear?
                 ' If ObjData(obj.ObjIndex).Log = 1 Then
                  '    Call LogDesarrollo(UserList(UserIndex).name & " tir� al piso " & obj.Amount & " " & ObjData(obj.ObjIndex).name)
              '    ElseIf obj.Amount = 1000 Then 'Es mucha cantidad?
               '   'Si no es de los prohibidos de loguear, lo logueamos.
                    '  If ObjData(obj.ObjIndex).NoLog <> 1 Then
                      '    Call LogDesarrollo(UserList(UserIndex).name & " tir� del piso " & obj.Amount & " " & ObjData(obj.ObjIndex).name)
                     ' End If
                 ' End If
            Else
              'Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
              Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call QuitarUserInvItem(UserIndex, slot, num)
            Call UpdateUserInv(False, UserIndex, slot)
        End If
    End If


End Sub

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer)
Dim Rango As Byte




MapData(Map, x, Y).ObjInfo.Amount = MapData(Map, x, Y).ObjInfo.Amount - num

If MapData(Map, x, Y).ObjInfo.Amount <= 0 Then

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
    MapData(Map, x, Y).ObjInfo.ObjIndex = 0
    MapData(Map, x, Y).ObjInfo.Amount = 0
    
    'Call Limpieza.Item_ListErase(Map, X, Y)
    
    Call modSendData.SendToAreaByPos(Map, x, Y, PrepareMessageObjectDelete(x, Y))
End If

End Sub

Sub MakeObj(ByRef obj As obj, ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer, Optional ByVal Limpiar As Boolean = True)
    Dim Color As Long
    Dim Rango As Byte

    If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
        If MapData(Map, x, Y).ObjInfo.ObjIndex = obj.ObjIndex Then
            MapData(Map, x, Y).ObjInfo.Amount = MapData(Map, x, Y).ObjInfo.Amount + obj.Amount
        Else
            MapData(Map, x, Y).ObjInfo.ObjIndex = obj.ObjIndex
            If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
                MapData(Map, x, Y).ObjInfo.Amount = ObjData(obj.ObjIndex).VidaUtil
            Else
                MapData(Map, x, Y).ObjInfo.Amount = obj.Amount
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
            Call modSendData.SendToAreaByPos(Map, x, Y, PrepareMessageObjectCreate(obj.ObjIndex, x, Y))
        
           ' If Limpiar Then
                'Call Limpieza.Item_ListAdd(Map, X, Y)
           ' End If
        End If
    
    End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim x As Integer
Dim Y As Integer
Dim slot As Byte

'�el user ya tiene un objeto del mismo tipo? ?????
If MiObj.ObjIndex = 12 Then
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.Amount

Else
    
    slot = 1
    Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And _
             UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
       slot = slot + 1
       If slot > UserList(UserIndex).CurrentInventorySlots Then
             Exit Do
       End If
    Loop
        
    'Sino busca un slot vacio
    If slot > UserList(UserIndex).CurrentInventorySlots Then
       slot = 1
       Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
           slot = slot + 1
           If slot > UserList(UserIndex).CurrentInventorySlots Then
               'Call WriteConsoleMsg(UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
               Call WriteLocaleMsg(UserIndex, "328", FontTypeNames.FONTTYPE_FIGHT)
               MeterItemEnInventario = False
               Exit Function
           End If
       Loop
       UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
    End If
        
    'Mete el objeto
    If UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
       'Menor que MAX_INV_OBJS
       UserList(UserIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex
       UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount + MiObj.Amount
    Else
       UserList(UserIndex).Invent.Object(slot).Amount = MAX_INVENTORY_OBJS
    End If
        
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, UserIndex, slot)
End If
WriteUpdateGold (UserIndex)
MeterItemEnInventario = True


Exit Function
Errhandler:

End Function
Function MeterItemEnInventarioDeNpc(ByVal NpcIndex As Integer, ByRef MiObj As obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim x As Integer
Dim Y As Integer
Dim slot As Byte

'�el user ya tiene un objeto del mismo tipo? ?????

    
    slot = 1
    Do Until Npclist(NpcIndex).Invent.Object(slot).ObjIndex = MiObj.ObjIndex And _
             Npclist(NpcIndex).Invent.Object(slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
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
Errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim obj As ObjData
Dim MiObj As obj

'�Hay algun obj?
If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    '�Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim x As Integer
        Dim Y As Integer
        Dim slot As Byte
        
        x = UserList(UserIndex).Pos.x
        Y = UserList(UserIndex).Pos.Y
        obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.ObjIndex
        
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            Else
            
                'Quitamos el objeto
                Call EraseObj(MapData(UserList(UserIndex).Pos.Map, x, Y).ObjInfo.Amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
                If Not UserList(UserIndex).flags.Privilegios And PlayerType.user Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
                
    
    
                If BusquedaTesoroActiva Then
                    If UserList(UserIndex).Pos.Map = TesoroNumMapa And UserList(UserIndex).Pos.x = TesoroX And UserList(UserIndex).Pos.Y = TesoroY Then
    
                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).name & " encontro el tesoro �Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
                        BusquedaTesoroActiva = False
                    End If
                End If
                
                
                If BusquedaRegaloActiva Then
                    If UserList(UserIndex).Pos.Map = RegaloNumMapa And UserList(UserIndex).Pos.x = RegaloX And UserList(UserIndex).Pos.Y = RegaloY Then
                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).name & " fue el valiente que encontro el gran item magico �Felicitaciones!", FontTypeNames.FONTTYPE_TALK))
                        BusquedaRegaloActiva = False
                    End If
                End If
                
                'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
                'Es un Objeto que tenemos que loguear?
                If ObjData(MiObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(UserList(UserIndex).name & " junt� del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
               ' ElseIf MiObj.Amount = 1000 Then 'Es mucha cantidad?
                  '  'Si no es de los prohibidos de loguear, lo logueamos.
                 '   'If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                       ' Call LogDesarrollo(UserList(UserIndex).name & " junt� del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name)
                   ' End If
                End If
                
            End If
    End If
Else
    If Not UserList(UserIndex).flags.UltimoMensaje = 261 Then
        Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 261
    End If
    
    'Call WriteConsoleMsg(UserIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal slot As Byte)
'Desequipa el item slot del inventario
Dim obj As ObjData


If (slot < LBound(UserList(UserIndex).Invent.Object)) Or (slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(slot).ObjIndex = 0 Then
    Exit Sub
End If

obj = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex)



Select Case obj.OBJType
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        UserList(UserIndex).Char.Arma_Aura = ""
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 1))

        
        
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            
            If UserList(UserIndex).flags.Montado = 0 Then
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        
        If obj.EfectoMagico = 14 Then
           UserList(UserIndex).flags.Da�oMagico = 0
        End If
            
        If obj.ResistenciaMagica > 0 Then
            UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - obj.ResistenciaMagica
            If UserList(UserIndex).flags.ResistenciaMagica < 0 Then UserList(UserIndex).flags.ResistenciaMagica = 0
        End If
    
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
        
    
   ' Case eOBJType.otAnillo
    '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
    '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
      ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
        
            
    Case eOBJType.OtHerramientas
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
        UserList(UserIndex).Invent.HerramientaEqpSlot = 0
        If UserList(UserIndex).flags.UsandoMacro = True Then
            Call WriteMacroTrabajoToggle(UserIndex, False)
        End If
        
        UserList(UserIndex).Char.WeaponAnim = NingunArma
            
         If UserList(UserIndex).flags.Montado = 0 Then
             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
         End If
    
       
    Case eOBJType.otmagicos
    
    
    
    
    
        Select Case obj.EfectoMagico
            Case 1 'Regenera Energia
                UserList(UserIndex).flags.RegeneracionSta = 0
            Case 2 'Modifica los Atributos
                UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                
                UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
               ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                Call WriteFYA(UserIndex)
            Case 3 'Modifica los skills
                UserList(UserIndex).Stats.UserSkills(obj.QueSkill) = UserList(UserIndex).Stats.UserSkills(obj.QueSkill) - obj.CuantoAumento
            Case 4 ' Regeneracion Vida
                UserList(UserIndex).flags.RegeneracionHP = 0
            Case 5 ' Regeneracion Mana
                UserList(UserIndex).flags.RegeneracionMana = 0
            Case 6 'Aumento Golpe
                UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit - obj.CuantoAumento
                UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - obj.CuantoAumento
            Case 7 '
                
            Case 9 ' Orbe Ignea
                UserList(UserIndex).flags.NoMagiaEfeceto = 0
            Case 10
                UserList(UserIndex).flags.incinera = 0
            Case 11
                UserList(UserIndex).flags.Paraliza = 0
            Case 12
                If UserList(UserIndex).flags.Muerto = 0 Then UserList(UserIndex).flags.CarroMineria = 0
                
            Case 14
                UserList(UserIndex).flags.Da�oMagico = 0
                
            Case 15 'Pendiete del Sacrificio
                 UserList(UserIndex).flags.PendienteDelSacrificio = 0
                 
            Case 16
                UserList(UserIndex).flags.NoPalabrasMagicas = 0
            Case 17 'Sortija de la verdad
                UserList(UserIndex).flags.NoDetectable = 0
            Case 18 ' Pendiente del Experto
                UserList(UserIndex).flags.PendienteDelExperto = 0
            Case 19
                UserList(UserIndex).flags.Envenena = 0
            Case 20 ' anillo de las sombras
                UserList(UserIndex).flags.AnilloOcultismo = 0
                
                
                
        End Select
        
        If obj.ResistenciaMagica > 0 Then
            UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - obj.ResistenciaMagica
            If UserList(UserIndex).flags.ResistenciaMagica < 0 Then UserList(UserIndex).flags.ResistenciaMagica = 0
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 5))
        UserList(UserIndex).Char.Otra_Aura = 0
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.MagicoObjIndex = 0
        UserList(UserIndex).Invent.MagicoSlot = 0
        
    Case eOBJType.otNUDILLOS
    
    
    'falta mandar animacion
    

            
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.NudilloObjIndex = 0
        UserList(UserIndex).Invent.NudilloSlot = 0
        
        UserList(UserIndex).Char.Arma_Aura = ""
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 1))
        
        UserList(UserIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        
        
    Case eOBJType.otArmadura
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
        UserList(UserIndex).Invent.ArmourEqpSlot = 0
        
        If UserList(UserIndex).flags.Navegando = 0 Then
            If UserList(UserIndex).flags.Montado = 0 Then
                Call DarCuerpoDesnudo(UserIndex)
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        End If
        
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 2))
    
    
    If obj.ResistenciaMagica > 0 Then
            UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - obj.ResistenciaMagica
            If UserList(UserIndex).flags.ResistenciaMagica < 0 Then UserList(UserIndex).flags.ResistenciaMagica = 0
        End If
        
    UserList(UserIndex).Char.Body_Aura = 0
    
    Case eOBJType.otCASCO
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.CascoEqpObjIndex = 0
        UserList(UserIndex).Invent.CascoEqpSlot = 0
        UserList(UserIndex).Char.Head_Aura = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 4))

        UserList(UserIndex).Char.CascoAnim = NingunCasco
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    
        
        If obj.ResistenciaMagica > 0 Then
            UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - obj.ResistenciaMagica
            If UserList(UserIndex).flags.ResistenciaMagica < 0 Then UserList(UserIndex).flags.ResistenciaMagica = 0
        End If
    
    Case eOBJType.otESCUDO
        UserList(UserIndex).Invent.Object(slot).Equipped = 0
        UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
        UserList(UserIndex).Invent.EscudoEqpSlot = 0
        UserList(UserIndex).Char.Escudo_Aura = 0
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, 0, True, 3))
        
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        If UserList(UserIndex).flags.Montado = 0 Then
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
        End If
        
        If obj.ResistenciaMagica > 0 Then
            UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica - obj.ResistenciaMagica
            If UserList(UserIndex).flags.ResistenciaMagica < 0 Then UserList(UserIndex).flags.ResistenciaMagica = 0
        End If
        
End Select


Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If EsGM(UserIndex) Then
    SexoPuedeUsarItem = True
    Exit Function
End If

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Status(UserIndex) = 3 Then
        FaccionPuedeUsarItem = esArmada(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Status(UserIndex) = 2 Then
        FaccionPuedeUsarItem = esCaos(UserIndex)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)
On Error GoTo Errhandler

Dim errordesc As String


'Equipa un item del inventario
Dim obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
obj = ObjData(ObjIndex)

If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
     Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If

If UserList(UserIndex).Stats.ELV < obj.MinELV And Not EsGM(UserIndex) Then
     Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If



Select Case obj.OBJType
    Case eOBJType.otWeapon

       If ClasePuedeUsarItem(UserIndex, ObjIndex, slot) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, slot)
                'Animacion por defecto

                    UserList(UserIndex).Char.WeaponAnim = NingunArma
                    If UserList(UserIndex).flags.Montado = 0 Then
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
            End If
            
            
            If UserList(UserIndex).Invent.NudilloObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudilloSlot)
            End If
            
            UserList(UserIndex).Invent.Object(slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = slot
            
            
            If obj.EfectoMagico = 14 Then
                UserList(UserIndex).flags.Da�oMagico = obj.CuantoAumento
            End If
            
            If obj.proyectil = 1 Then 'Si es un arco, desequipa el escudo.
            
                'If UserList(UserIndex).Invent.EscudoEqpObjIndex = 404 Or UserList(UserIndex).Invent.EscudoEqpObjIndex = 1007 Or UserList(UserIndex).Invent.EscudoEqpObjIndex = 1358 Then
                If UserList(UserIndex).Invent.EscudoEqpObjIndex = 1700 Or UserList(UserIndex).Invent.EscudoEqpObjIndex = 1730 Or UserList(UserIndex).Invent.EscudoEqpObjIndex = 1724 Or UserList(UserIndex).Invent.EscudoEqpObjIndex = 1717 Or UserList(UserIndex).Invent.EscudoEqpObjIndex = 1699 Then
                
                Else
                    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                         Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                         Call WriteConsoleMsg(UserIndex, "No podes tirar flechas si ten�s un escudo equipado. Tu escudo fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                     End If
                End If
            End If
            
            errordesc = "Arma"
            If obj.ResistenciaMagica > 0 Then
                UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + obj.ResistenciaMagica
            End If
            
            'Sonido
            If obj.SndAura = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            End If
            
           If obj.CreaGRH <> "" Then
                UserList(UserIndex).Char.Arma_Aura = obj.CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 1))
            End If
       

                
                
                If UserList(UserIndex).flags.Montado = 0 Then
                    If UserList(UserIndex).flags.Navegando = 0 Then
                        UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                End If
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    

      
      
          Case eOBJType.OtHerramientas
             'Si esta equipado lo quita
             If UserList(UserIndex).Invent.Object(slot).Equipped Then
                 'Quitamos del inv el item
                 Call Desequipar(UserIndex, slot)
                 Exit Sub
             End If
             If obj.MinSkill <> 0 Then
                  If UserList(UserIndex).Stats.UserSkills(obj.QueSkill) < obj.MinSkill Then
                      Call WriteConsoleMsg(UserIndex, "Para podes usar " & obj.name & " necesitas al menos " & obj.MinSkill & " puntos en " & SkillsNames(obj.QueSkill) & ".", FontTypeNames.FONTTYPE_INFOIAO)
                      Exit Sub
                  End If
              End If
             'Quitamos el elemento anterior
             If UserList(UserIndex).Invent.HerramientaEqpObjIndex > 0 Then
                 Call Desequipar(UserIndex, UserList(UserIndex).Invent.HerramientaEqpSlot)
             End If
             
              If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                     Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
              End If
             
             UserList(UserIndex).Invent.Object(slot).Equipped = 1
             UserList(UserIndex).Invent.HerramientaEqpObjIndex = ObjIndex
             UserList(UserIndex).Invent.HerramientaEqpSlot = slot
             
            If UserList(UserIndex).flags.Montado = 0 Then
              If UserList(UserIndex).flags.Navegando = 0 Then
                  UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
                  Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
              End If
            End If
       
       
    Case eOBJType.otmagicos
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
           'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        errordesc = "Magico"
        
                        'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MagicoObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MagicoSlot)
                End If
        
                UserList(UserIndex).Invent.Object(slot).Equipped = 1
                UserList(UserIndex).Invent.MagicoObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
                UserList(UserIndex).Invent.MagicoSlot = slot
                
               ' Debug.Print "magico" & obj.EfectoMagico
        Select Case obj.EfectoMagico
            Case 1 ' Regenera Stamina
                UserList(UserIndex).flags.RegeneracionSta = 1

            Case 2 'Modif la fuerza, agilidad, carisma, etc
               ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo)
                
                UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
                
                
                UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento
                If UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = MAXATRIBUTOS
                
                Call WriteFYA(UserIndex)
            Case 3 'Modifica los skills
            
                UserList(UserIndex).Stats.UserSkills(obj.QueSkill) = UserList(UserIndex).Stats.UserSkills(obj.QueSkill) + obj.CuantoAumento
            Case 4
                UserList(UserIndex).flags.RegeneracionHP = 1
            Case 5
                UserList(UserIndex).flags.RegeneracionMana = 1
            Case 6
                'Call WriteConsoleMsg(UserIndex, "Item, temporalmente deshabilitado.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit + obj.CuantoAumento
                UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT + obj.CuantoAumento
            Case 9
                UserList(UserIndex).flags.NoMagiaEfeceto = 1
            Case 10
                UserList(UserIndex).flags.incinera = 1
            Case 11
                UserList(UserIndex).flags.Paraliza = 1
            Case 12
                 UserList(UserIndex).flags.CarroMineria = 1
                
            Case 14
                UserList(UserIndex).flags.Da�oMagico = obj.CuantoAumento
                
            Case 15 'Pendiete del Sacrificio
                 UserList(UserIndex).flags.PendienteDelSacrificio = 1
            Case 16
                UserList(UserIndex).flags.NoPalabrasMagicas = 1
            Case 17
                UserList(UserIndex).flags.NoDetectable = 1
                   
            Case 18 ' Pendiente del Experto
                UserList(UserIndex).flags.PendienteDelExperto = 1
            Case 19
                UserList(UserIndex).flags.Envenena = 1
            Case 20 'Anillo ocultismo
                UserList(UserIndex).flags.AnilloOcultismo = 1


    
        End Select
            errordesc = "Magico"
            If obj.ResistenciaMagica > 0 Then
                UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + obj.ResistenciaMagica
            End If
            
             'Sonido
            If obj.SndAura <> 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            End If
            
            If obj.CreaGRH <> "" Then
            
                UserList(UserIndex).Char.Otra_Aura = obj.CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Otra_Aura, False, 5))
            End If
    
        
            'Call WriteUpdateExp(UserIndex)
           ' Call CheckUserLevel(UserIndex)
            
    Case eOBJType.otNUDILLOS
    
    
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                   'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                 If Not ClasePuedeUsarItem(UserIndex, ObjIndex, slot) Then
                    Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                 End If
            
                 
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                End If

                If UserList(UserIndex).Invent.Object(slot).Equipped Then
                    Call Desequipar(UserIndex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.NudilloObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudilloSlot)
                End If
        
                UserList(UserIndex).Invent.Object(slot).Equipped = 1
                UserList(UserIndex).Invent.NudilloObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
                UserList(UserIndex).Invent.NudilloSlot = slot
    
        
                'Falta enviar anim
                
                
                If UserList(UserIndex).flags.Montado = 0 Then
                    If UserList(UserIndex).flags.Navegando = 0 Then
                        UserList(UserIndex).Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                End If
            
                 If obj.SndAura = 0 Then
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                 Else
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                 End If
                 
                If obj.CreaGRH <> "" Then
                     UserList(UserIndex).Char.Arma_Aura = obj.CreaGRH
                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Arma_Aura, False, 1))
                 End If
                 


        
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(slot).Equipped Then
                    'Quitamos del inv el item
                                        
                    Call Desequipar(UserIndex, slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = slot
                
       Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex, slot) And _
           SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) And _
           CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) And _
           FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then
           
           'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(slot).Equipped Then
                Call Desequipar(UserIndex, slot)
                If UserList(UserIndex).flags.Navegando = 0 Then
                    If UserList(UserIndex).flags.Montado = 0 Then
                        Call DarCuerpoDesnudo(UserIndex)
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                End If
                Exit Sub
            End If
            'Quita el anterior
            If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            errordesc = "Armadura 2"
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
            errordesc = "Armadura 3"
            End If
            

                
            
            If obj.ResistenciaMagica > 0 Then
                UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + obj.ResistenciaMagica
            End If
  
            'Lo equipa
            If obj.CreaGRH <> "" Then
                 UserList(UserIndex).Char.Body_Aura = obj.CreaGRH
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Body_Aura, False, 2))
            End If
            
            UserList(UserIndex).Invent.Object(slot).Equipped = 1
            UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
            UserList(UserIndex).Invent.ArmourEqpSlot = slot
                            
            If UserList(UserIndex).flags.Montado = 0 Then
                If UserList(UserIndex).flags.Navegando = 0 Then
        
                If UserList(UserIndex).raza = Enano Or UserList(UserIndex).raza = Gnomo Then
                    If obj.RopajeBajo <> 0 Then
                    UserList(UserIndex).Char.Body = obj.RopajeBajo
                    Else
                    UserList(UserIndex).Char.Body = obj.Ropaje
                    End If
                    
                    
                    
                Else
                    UserList(UserIndex).Char.Body = obj.Ropaje
                End If
                
                If UserList(UserIndex).Char.Body = 0 And EsGM(UserIndex) Then
                    UserList(UserIndex).Char.Body = obj.RopajeBajo
                End If
                
                
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                UserList(UserIndex).flags.Desnudo = 0
            
            End If
        End If
        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otCASCO
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex, slot) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(slot).Equipped Then
                Call Desequipar(UserIndex, slot)
                
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
            End If
            
            If obj.ResistenciaMagica > 0 Then
                UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + obj.ResistenciaMagica
            End If
            
            errordesc = "Casco"
            'Lo equipa
            If obj.CreaGRH <> "" Then
                 UserList(UserIndex).Char.Head_Aura = obj.CreaGRH
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Head_Aura, False, 4))
             End If
            
            UserList(UserIndex).Invent.Object(slot).Equipped = 1
            UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
            UserList(UserIndex).Invent.CascoEqpSlot = slot
            
            If UserList(UserIndex).flags.Navegando = 0 Then
                UserList(UserIndex).Char.CascoAnim = obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otESCUDO
         If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex, slot) And _
             FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(UserIndex).Invent.Object(slot).Equipped Then
                 Call Desequipar(UserIndex, slot)
                
                 
                     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                    If UserList(UserIndex).flags.Montado = 0 Then
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                 Exit Sub
             End If
             
             
     
             'Quita el anterior
             If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             If UserList(UserIndex).Invent.Object(slot).ObjIndex = 1700 Or UserList(UserIndex).Invent.Object(slot).ObjIndex = 1730 Or UserList(UserIndex).Invent.Object(slot).ObjIndex = 1724 Or UserList(UserIndex).Invent.Object(slot).ObjIndex = 1717 Or UserList(UserIndex).Invent.Object(slot).ObjIndex = 1699 Then
             
             Else
                 If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                     If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                            Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                             Call WriteConsoleMsg(UserIndex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                End If
            End If
            
            errordesc = "Escudo"
            If obj.ResistenciaMagica > 0 Then
                UserList(UserIndex).flags.ResistenciaMagica = UserList(UserIndex).flags.ResistenciaMagica + obj.ResistenciaMagica
            End If
             
            If obj.CreaGRH <> "" Then
                 UserList(UserIndex).Char.Escudo_Aura = obj.CreaGRH
                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).Char.Escudo_Aura, False, 3))
             End If
             UserList(UserIndex).Invent.Object(slot).Equipped = 1
             UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
             UserList(UserIndex).Invent.EscudoEqpSlot = slot
             
             
                 
                If UserList(UserIndex).flags.Navegando = 0 Then
                    If UserList(UserIndex).flags.Montado = 0 Then
                        UserList(UserIndex).Char.ShieldAnim = obj.ShieldAnim
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    End If
                End If
         Else
             Call WriteConsoleMsg(UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, slot)



Exit Sub
Errhandler:
Debug.Print errordesc
Call LogError("EquiparInvItem Slot:" & slot & " - Error: " & Err.Number & " - Error Description : " & Err.description & "- " & errordesc)
End Sub

Public Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler


If EsGM(UserIndex) Then
    CheckRazaUsaRopa = True
    Exit Function
End If

Select Case UserList(UserIndex).raza

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
        If ObjData(ItemIndex).RopajeBajo > 0 Then
            CheckRazaUsaRopa = True
            Exit Function
        End If
    
        
    Case eRaza.Enano
        If ObjData(ItemIndex).RopajeBajo > 0 Then
            CheckRazaUsaRopa = True
            Exit Function
        End If
    
End Select

CheckRazaUsaRopa = False
    


Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function
Public Function CheckRazaTipo(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

If EsGM(UserIndex) Then

CheckRazaTipo = True
Exit Function
End If

Select Case ObjData(ItemIndex).RazaTipo
    Case 0
        CheckRazaTipo = True
    Case 1
        If UserList(UserIndex).raza = eRaza.Elfo Then
        CheckRazaTipo = True
        Exit Function
        End If
        
        If UserList(UserIndex).raza = eRaza.Drow Then
        CheckRazaTipo = True
        Exit Function
        End If
        
        If UserList(UserIndex).raza = eRaza.Humano Then
        CheckRazaTipo = True
        Exit Function
        End If
    Case 2
        If UserList(UserIndex).raza = eRaza.Gnomo Then CheckRazaTipo = True
        If UserList(UserIndex).raza = eRaza.Enano Then CheckRazaTipo = True
        Exit Function
    Case 3
        If UserList(UserIndex).raza = eRaza.Orco Then CheckRazaTipo = True
        Exit Function
    
    
End Select



  

Exit Function
Errhandler:
    Call LogError("Error CheckRazaTipo ItemIndex:" & ItemIndex)

End Function
Public Function CheckClaseTipo(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler


If EsGM(UserIndex) Then

CheckClaseTipo = True
Exit Function
End If



Select Case ObjData(ItemIndex).ClaseTipo
    Case 0
        CheckClaseTipo = True
        Exit Function
    Case 2
        If UserList(UserIndex).clase = eClass.Mage Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Druid Then CheckClaseTipo = True
        Exit Function

    Case 1
        If UserList(UserIndex).clase = eClass.Warrior Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Assasin Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Bard Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Cleric Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Paladin Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Trabajador Then CheckClaseTipo = True
        If UserList(UserIndex).clase = eClass.Hunter Then CheckClaseTipo = True
        Exit Function

End Select
  

Exit Function
Errhandler:
    Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal slot As Byte)


On Error GoTo hErr
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legi�n.
'24/01/2007 Pablo (ToxicWaste) - Utilizaci�n nueva de Barco en lvl 20 por clase Pirata y Pescador.
'*************************************************

Dim obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As obj

If UserList(UserIndex).Invent.Object(slot).Amount = 0 Then Exit Sub

obj = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex)

If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
    Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < obj.MinELV Then
     Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If

If obj.OBJType = eOBJType.otWeapon Then
    If obj.proyectil = 1 Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
UserList(UserIndex).flags.TargetObjInvIndex = ObjIndex
UserList(UserIndex).flags.TargetObjInvSlot = slot

Select Case obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
           ' Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + obj.MinHam
        If UserList(UserIndex).Stats.MinHam > UserList(UserIndex).Stats.MaxHam Then _
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        UserList(UserIndex).flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.MORFAR_MANZANA, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, slot, 1)
        
        Call UpdateUserInv(False, UserIndex, slot)

    Case eOBJType.otGuita
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
           ' Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + UserList(UserIndex).Invent.Object(slot).Amount
        UserList(UserIndex).Invent.Object(slot).Amount = 0
        UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
        UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
           ' Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not UserList(UserIndex).Stats.MinSta > 0 Then
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        Else
            If UserList(UserIndex).flags.TargetObj = Le�a Then
                If UserList(UserIndex).Invent.Object(slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(UserIndex).flags.TargetObjMap, _
                         UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY, UserIndex)
                End If
            End If
        End If
        
        'REVISAR LADDER
        'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
        If UserList(UserIndex).Invent.Object(slot).Equipped = 0 Then
            'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
            'Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        
    Case eOBJType.OtHerramientas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not UserList(UserIndex).Stats.MinSta > 0 Then
            Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
        If UserList(UserIndex).Invent.Object(slot).Equipped = 0 Then
            'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "376", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

    Select Case ObjIndex
            Case CA�A_PESCA, RED_PESCA, HACHA_LE�ADOR, TIJERAS, PIQUETE_MINERO, CA�A_PESCA_DORADA, TIJERAS_DORADAS, HACHA_LE�ADOR_DORADA, PIQUETE_MINERO_DORADA
                Call WriteWorkRequestTarget(UserIndex, eSkill.Recoleccion)
            Case MARTILLO_HERRERO
                Call WriteConsoleMsg(UserIndex, "Debes hacer click derecho sobre el yunke.", FontTypeNames.FONTTYPE_INFOIAO)
               ' Call WriteWorkRequestTarget(UserIndex, eSkill.Manualidades)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(UserIndex)
                Call WriteShowCarpenterForm(UserIndex)
            Case OLLA_ALQUIMIA
                Call EnivarObjConstruiblesAlquimia(UserIndex)
                Call WriteShowAlquimiaForm(UserIndex)
            Case COSTURERO
                Call EnivarObjConstruiblesSastre(UserIndex)
                Call WriteShowSastreForm(UserIndex)
        End Select
        
    
    Case eOBJType.otPociones
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(UserIndex).flags.TomoPocion = True
        UserList(UserIndex).flags.TipoPocion = obj.TipoPocion
                
        Dim CabezaFinal As Integer
        Dim CabezaActual As Integer
        Select Case UserList(UserIndex).flags.TipoPocion
        
        
            Case 1 'Modif la agilidad
                UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                'If UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Agilidad)
                Call WriteFYA(UserIndex)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, slot, 1)
                If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If
        
            Case 2 'Modif la fuerza
                UserList(UserIndex).flags.DuracionEfecto = obj.DuracionEfecto
        
                'Usa el item
                UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                'If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(UserIndex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, slot, 1)
                If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If
                Call WriteFYA(UserIndex)
            Case 3 'Pocion roja, restaura HP
                
                
                'Usa el item
                UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)
                If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then _
                    UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, slot, 1)
                If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If

            
            Case 4 'Pocion azul, restaura MANA
            
                Dim porcentajeRec As Byte
                
                Select Case UserList(UserIndex).clase
                    Case Paladin, Assasin
                        porcentajeRec = obj.Porcentaje * 1.4
                    Case Else
                        porcentajeRec = obj.Porcentaje
                End Select
                
                
                'Usa el item
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Porcentaje(UserList(UserIndex).Stats.MaxMAN, porcentajeRec)
                If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, slot, 1)
                If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If
                
            Case 5 ' Pocion violeta
                If UserList(UserIndex).flags.Envenenado > 0 Then
                    UserList(UserIndex).flags.Envenenado = 0
                    Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                End If
                Else
                    Call WriteConsoleMsg(UserIndex, "�No te encuentras envenenado!", FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case 6  ' Remueve Par�lisis
                    If UserList(UserIndex).flags.Paralizado = 1 Or UserList(UserIndex).flags.Inmovilizado = 1 Then
                        If UserList(UserIndex).flags.Paralizado = 1 Then
                            UserList(UserIndex).flags.Paralizado = 0
                            Call WriteParalizeOK(UserIndex)
                        End If
                        
                        If UserList(UserIndex).flags.Inmovilizado = 1 Then
                            UserList(UserIndex).Counters.Inmovilizado = 0
                            UserList(UserIndex).flags.Inmovilizado = 0
                            Call WriteInmovilizaOK(UserIndex)
                        End If
                        
                        Call FlushBuffer(UserIndex)
                        
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(255, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        End If
                            Call WriteConsoleMsg(UserIndex, "Te has removido la paralizis.", FontTypeNames.FONTTYPE_INFOIAO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No estas paralizado.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                    
                    
                
                
            Case 7  ' Pocion Naranja
                    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)
                    If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then _
                        UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
                    
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                            
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    End If
            Case 8  ' Pocion cambio cara
                Select Case UserList(UserIndex).genero
                    Case eGenero.Hombre
                        Select Case UserList(UserIndex).raza
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
                            Select Case UserList(UserIndex).raza
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
            
                    UserList(UserIndex).Char.Head = CabezaFinal
                    UserList(UserIndex).OrigChar.Head = CabezaFinal
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, CabezaFinal, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    'Quitamos del inv el item
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 102, 0))
                    If CabezaActual <> CabezaFinal Then
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                    Else
                        Call WriteConsoleMsg(UserIndex, "�Rayos! Te toc� la misma cabeza, item no consumido. Tienes otra oportunidad.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    
            Case 9  ' Pocion sexo
            

    
                Select Case UserList(UserIndex).genero
                    Case eGenero.Hombre
                         UserList(UserIndex).genero = eGenero.Mujer
                    
                    Case eGenero.Mujer
                        UserList(UserIndex).genero = eGenero.Hombre
                    
                    End Select
                    
            
                     Select Case UserList(UserIndex).genero
                    Case eGenero.Hombre
                        Select Case UserList(UserIndex).raza
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
                            Select Case UserList(UserIndex).raza
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
            
                    UserList(UserIndex).Char.Head = CabezaFinal
                    UserList(UserIndex).OrigChar.Head = CabezaFinal
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, CabezaFinal, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    'Quitamos del inv el item
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 102, 0))
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    End If
                
            Case 10  ' Invisibilidad
            
            
                    If UserList(UserIndex).flags.invisible = 0 Then
                        UserList(UserIndex).flags.invisible = 1
                        UserList(UserIndex).Counters.Invisibilidad = obj.DuracionEfecto
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, True))
                        Call WriteContadores(UserIndex)
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                            
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("123", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        End If
                        Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras...", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ya estas invisible.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        Exit Sub
                    End If
                    
                    
            Case 11  ' Experiencia
                        Dim HR As Integer
                        Dim MS As Integer
                        Dim SS As Integer
                        Dim secs As Integer
                    If UserList(UserIndex).flags.ScrollExp = 1 Then
                        UserList(UserIndex).flags.ScrollExp = obj.CuantoAumento
                        UserList(UserIndex).Counters.ScrollExperiencia = obj.DuracionEfecto
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        
                        
                        secs = obj.DuracionEfecto
                        HR = secs \ 3600
                        MS = (secs Mod 3600) \ 60
                        SS = (secs Mod 3600) Mod 60
                        If SS > 9 Then
                        Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                        Else
                        Call WriteConsoleMsg(UserIndex, "Tu scroll de experiencia ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                        Exit Sub
                    End If
                    Call WriteContadores(UserIndex)
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    End If
                Case 12  ' Oro
            
            
                    If UserList(UserIndex).flags.ScrollOro = 1 Then
                        UserList(UserIndex).flags.ScrollOro = obj.CuantoAumento
                        UserList(UserIndex).Counters.ScrollOro = obj.DuracionEfecto
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                        secs = obj.DuracionEfecto
                        HR = secs \ 3600
                        MS = (secs Mod 3600) \ 60
                        SS = (secs Mod 3600) Mod 60
                        If SS > 9 Then
                        Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                        Else
                        Call WriteConsoleMsg(UserIndex, "Tu scroll de oro ha comenzado. Este beneficio durara: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_DONADOR)
                        End If
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "Solo podes usar un scroll a la vez.", FontTypeNames.FONTTYPE_New_DONADOR)
                        Exit Sub
                    End If
                    Call WriteContadores(UserIndex)
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    End If
                Case 13
                
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    UserList(UserIndex).flags.Envenenado = 0
                    UserList(UserIndex).flags.Incinerado = 0
                    
                    If UserList(UserIndex).flags.Inmovilizado = 1 Then
                        UserList(UserIndex).Counters.Inmovilizado = 0
                        UserList(UserIndex).flags.Inmovilizado = 0
                        Call WriteInmovilizaOK(UserIndex)
                        Call FlushBuffer(UserIndex)
                    End If
                    
                    If UserList(UserIndex).flags.Paralizado = 1 Then
                        UserList(UserIndex).flags.Paralizado = 0
                        Call WriteParalizeOK(UserIndex)
                        Call FlushBuffer(UserIndex)
                    End If
                    
                    If UserList(UserIndex).flags.Ceguera = 1 Then
                        UserList(UserIndex).flags.Ceguera = 0
                        Call WriteBlindNoMore(UserIndex)
                        Call FlushBuffer(UserIndex)
                    End If
                    
                    
                    If UserList(UserIndex).flags.Maldicion = 1 Then
                        UserList(UserIndex).flags.Maldicion = 0
                        UserList(UserIndex).Counters.Maldicion = 0
                    End If
                    
                    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta
                    UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
                    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                    UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                    UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
                    
                    UserList(UserIndex).flags.Hambre = 0
                    UserList(UserIndex).flags.Sed = 0
                    
                    Call WriteUpdateHungerAndThirst(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Donador> Te sentis sano y lleno.", FontTypeNames.FONTTYPE_WARNING)
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    End If
                Case 14
                
                    If UserList(UserIndex).flags.BattleModo = 1 Then
                        Call WriteConsoleMsg(UserIndex, "No podes usarlo aqu�.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    
                    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = CARCEL Then
                        Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    
                    Dim Map As Integer
                    Dim x As Byte
                    Dim Y As Byte
                    Dim DeDonde As WorldPos
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    
            
                    Select Case UserList(UserIndex).Hogar
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
                    x = DeDonde.x
                    Y = DeDonde.Y
                    
                    Call FindLegalPos(UserIndex, Map, x, Y)
                    Call WarpUserChar(UserIndex, Map, x, Y, True)
                    Call WriteConsoleMsg(UserIndex, "Ya estas a salvo...", FontTypeNames.FONTTYPE_WARNING)
                    If obj.Snd1 <> 0 Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        
                    Else
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                    End If
                Case 15  ' Aliento de sirena
                
                        
                        If UserList(UserIndex).Counters.Oxigeno >= 3540 Then
                        
                        
                             Call WriteConsoleMsg(UserIndex, "No podes acumular m�s de 59 minutos de oxigeno.", FontTypeNames.FONTTYPE_INFOIAO)
                            secs = UserList(UserIndex).Counters.Oxigeno
                            HR = secs \ 3600
                            MS = (secs Mod 3600) \ 60
                            SS = (secs Mod 3600) Mod 60
                            If SS > 9 Then
                                Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":" & SS & " segundos.", FontTypeNames.FONTTYPE_New_Blanco)
                            Else
                            Call WriteConsoleMsg(UserIndex, "Tu reserva de oxigeno es de " & HR & ":" & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_New_Blanco)
                            End If
                        Else
                            
                            UserList(UserIndex).Counters.Oxigeno = UserList(UserIndex).Counters.Oxigeno + obj.DuracionEfecto
                            Call QuitarUserInvItem(UserIndex, slot, 1)

                            
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
                            
                            UserList(UserIndex).flags.Ahogandose = 0
                        Call WriteOxigeno(UserIndex)
                            
                        Call WriteContadores(UserIndex)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                            
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        End If
                    End If
                Case 16 ' Divorcio
                    If UserList(UserIndex).flags.Casado = 1 Then
                        Dim tUser As Integer
                        'UserList(UserIndex).flags.Pareja
                        tUser = NameIndex(UserList(UserIndex).flags.Pareja)
                         Call QuitarUserInvItem(UserIndex, slot, 1)
                        
                        If tUser <= 0 Then
                            Dim FileUser As String
                            FileUser = CharPath & UCase$(UserList(UserIndex).flags.Pareja) & ".chr"
                            Call WriteVar(FileUser, "FLAGS", "CASADO", 0)
                            Call WriteVar(FileUser, "FLAGS", "PAREJA", "")
                            UserList(UserIndex).flags.Casado = 0
                            UserList(UserIndex).flags.Pareja = ""
                            Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", UserList(UserIndex).name & " se ha divorciado de ti.")
                        

                        Else
                            UserList(tUser).flags.Casado = 0
                            UserList(tUser).flags.Pareja = ""
                            UserList(UserIndex).flags.Casado = 0
                            UserList(UserIndex).flags.Pareja = ""
                            Call WriteConsoleMsg(UserIndex, "Te has divorciado.", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteConsoleMsg(tUser, UserList(UserIndex).name & " se ha divorciado de ti.", FontTypeNames.FONTTYPE_INFOIAO)
                            
                        End If
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                            
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        End If
                    
                    Else
                        Call WriteConsoleMsg(UserIndex, "No estas casado.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                Case 17 'Cara legendaria
                Select Case UserList(UserIndex).genero
                    Case eGenero.Hombre
                        Select Case UserList(UserIndex).raza
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
                            Select Case UserList(UserIndex).raza
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
                    CabezaActual = UserList(UserIndex).OrigChar.Head
                        
                    UserList(UserIndex).Char.Head = CabezaFinal
                    UserList(UserIndex).OrigChar.Head = CabezaFinal
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, CabezaFinal, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    'Quitamos del inv el item
                    If CabezaActual <> CabezaFinal Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 102, 0))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                    Else
                        Call WriteConsoleMsg(UserIndex, "�Rayos! No pude asignarte una cabeza nueva, item no consumido. �Proba de nuevo!", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                Case 18  ' tan solo crea una particula por determinado tiempo
                        Dim Particula As Integer
                        Dim Tiempo As Long
                        Dim ParticulaPermanente As Byte
                        Dim sobrechar As Byte
                        If obj.CreaParticula <> "" Then
                            Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
                            Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
                            ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
                            sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                            
                            If ParticulaPermanente = 1 Then
                            UserList(UserIndex).Char.ParticulaFx = Particula
                            UserList(UserIndex).Char.loops = Tiempo
                            End If
                            
                            If sobrechar = 1 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, Particula, Tiempo))
                            Else
                            
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, Particula, Tiempo, False))
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y, Particula, Tiempo))
                            End If
                        End If
                        
                        If obj.CreaFX <> 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                            
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, obj.CreaFX, 0))
                           ' PrepareMessageCreateFX
                        End If
                        
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                        End If
                        
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                Case 19 ' Reseteo de skill
                    Dim S As Byte
                
                    If UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) >= 80 Then
                        Call WriteConsoleMsg(UserIndex, "Has fundado un clan, no podes resetar tus skills. ", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                    
                    For S = 1 To NUMSKILLS
                        UserList(UserIndex).Stats.UserSkills(S) = 0
                    Next S
                    
                    Dim SkillLibres As Integer
                    
                    SkillLibres = 5
                    SkillLibres = SkillLibres + (5 * UserList(UserIndex).Stats.ELV)
                    
                     
                    UserList(UserIndex).Stats.SkillPts = SkillLibres
                    Call WriteLevelUp(UserIndex, UserList(UserIndex).Stats.SkillPts)
                    
                    Call WriteConsoleMsg(UserIndex, "Tus skills han sido reseteados.", FontTypeNames.FONTTYPE_INFOIAO)
                     Call QuitarUserInvItem(UserIndex, slot, 1)
                Case 20
                
                    If UserList(UserIndex).Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
                        UserList(UserIndex).Stats.InventLevel = UserList(UserIndex).Stats.InventLevel + 1
                        UserList(UserIndex).CurrentInventorySlots = getMaxInventorySlots(UserIndex)
                        Call WriteInventoryUnlockSlots(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Has aumentado el espacio de tu inventario!", FontTypeNames.FONTTYPE_INFO)
                        Call QuitarUserInvItem(UserIndex, slot, 1)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ya has desbloqueado todos los casilleros disponibles.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                End Select
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, slot)

     Case eOBJType.otBebidas
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, slot, 1)
        
        If obj.Snd1 <> 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        End If
        
        Call UpdateUserInv(False, UserIndex, slot)
        
        
    Case eOBJType.OtCofre
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        
        

        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, slot, 1)
        Call UpdateUserInv(False, UserIndex, slot)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha abierto un " & obj.name & " y obtuvo...", FontTypeNames.FONTTYPE_New_DONADOR))
        
        If obj.Snd1 <> 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
        End If
        
        If obj.CreaFX <> 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, obj.CreaFX, 0))
        End If
        
        
        
        
        Dim i As Byte
        If obj.Subtipo = 1 Then

        For i = 1 To obj.CantItem
            If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then _
                Call TirarItemAlPiso(UserList(UserIndex).Pos, obj.Item(i))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
        Next i
        
        Else
        
            For i = 1 To obj.CantEntrega
                Dim indexobj As Byte
                
                indexobj = RandomNumber(1, obj.CantItem)
            
            
                Dim Index As obj
                Index.ObjIndex = obj.Item(indexobj).ObjIndex
                Index.Amount = obj.Item(indexobj).Amount
                If Not MeterItemEnInventario(UserIndex, Index) Then _
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, Index)
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).name & " (" & Index.Amount & ")", FontTypeNames.FONTTYPE_INFOBOLD))
            Next i
        End If
        
           
        
    
    Case eOBJType.otLlaves
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
        '�El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '�Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '�Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = obj.clave Then
         
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = obj.clave Then
                        MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not HayAgua(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY) Then
            Call WriteConsoleMsg(UserIndex, "No hay agua all�.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, slot)
    
    Case eOBJType.otBotellaLlena
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
           ' Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU + obj.MinSed
        If UserList(UserIndex).Stats.MinAGU > UserList(UserIndex).Stats.MaxAGU Then _
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MaxAGU
        UserList(UserIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.Amount = 1
        MiObj.ObjIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, slot)
    
    Case eOBJType.otPergaminos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
           ' Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Call LogError(UserList(UserIndex).Name & " intento aprender el hechizo " & ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex)
        
        
        If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(slot).ObjIndex, slot) Then
                'If UserList(UserIndex).Stats.MaxMAN > 0 Then
                    If UserList(UserIndex).flags.Hambre = 0 And _
                        UserList(UserIndex).flags.Sed = 0 Then
                        Call AgregarHechizo(UserIndex, slot)
                        Call UpdateUserInv(False, UserIndex, slot)
                       ' Call LogError(UserList(UserIndex).Name & " lo aprendio.")
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
                    End If
               ' Else
                '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_WARNING)
                'End If
        Else
             
                Call WriteConsoleMsg(UserIndex, "Por mas que lo intentas, no pod�s comprender el manuescrito.", FontTypeNames.FONTTYPE_INFO)
   
        End If
            
                
            
        
        
        
        
    Case eOBJType.otMinerales
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
             'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If obj.Real Then '�Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aqu�. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf obj.Caos Then '�Es el Cuerno Legi�n?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No hay Peligro aqu�. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legi�n Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
       
    Case eOBJType.otBarcos
        'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(UserIndex).Stats.ELV < 25 Then
                    Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
        End If
        
    'If obj.Subtipo = 0 Then
        If UserList(UserIndex).flags.Navegando = 0 Then
            If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x - 1, UserList(UserIndex).Pos.Y, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y - 1, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x + 1, UserList(UserIndex).Pos.Y, True, False) _
                    Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y + 1, True, False)) _
                    And UserList(UserIndex).flags.Navegando = 0) _
                    Or UserList(UserIndex).flags.Navegando = 1 Then
                Call DoNavega(UserIndex, obj, slot)
            Else
                Call WriteConsoleMsg(UserIndex, "�Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
            End If
        Else 'Ladder 10-02-2010
            If UserList(UserIndex).Invent.BarcoObjIndex <> UserList(UserIndex).Invent.Object(slot).ObjIndex Then
                Call DoReNavega(UserIndex, obj, slot)
            Else
                If ((LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x - 1, UserList(UserIndex).Pos.Y, False, True) _
                    Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y - 1, False, True) _
                    Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x + 1, UserList(UserIndex).Pos.Y, False, True) _
                    Or LegalPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y + 1, False, True)) _
                    And UserList(UserIndex).flags.Navegando = 1) _
                    Or UserList(UserIndex).flags.Navegando = 0 Then
                    Call DoNavega(UserIndex, obj, slot)
                Else
                    Call WriteConsoleMsg(UserIndex, "�Debes aproximarte a la costa para dejar la barca!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    'Else
    
    
   ' End If
    
        
        
   ' Case eOBJType.otTrajeDeBa�o
        
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
          '  Call WriteConsoleMsg(UserIndex, "�Te hago nadar!", FontTypeNames.FONTTYPE_INFO)
          '  Call DoNadar(UserIndex, ObjData(UserList(UserIndex).Invent.BarcoObjIndex), 0)
           ' Call DoNavega(UserIndex, ObjData(UserList(UserIndex).Invent.BarcoObjIndex), UserList(UserIndex).Invent.BarcoSlot)
            
            'UserList(UserIndex).flags.Nadando = 1
            
       ' Else
         '   Call WriteConsoleMsg(UserIndex, "�No podes nadar!", FontTypeNames.FONTTYPE_INFO)
       ' End If
          

        
    Case eOBJType.otMonturas
    'Verifica todo lo que requiere la montura
    
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "�Estas muerto! Los fantasmas no pueden montar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
            
        If UserList(UserIndex).flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "Debes dejar de navegar para poder montart�.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Meditando = True Then
            Call WriteConsoleMsg(UserIndex, "No pod�s subirte a la montura si estas meditando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapInfo(UserList(UserIndex).Pos.Map).zone = "DUNGEON" Then
                Call WriteConsoleMsg(UserIndex, "No podes cabalgar dentro de un dungeon.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call DoMontar(UserIndex, obj, slot)
    Case eOBJType.OtDonador
        Select Case obj.Subtipo
            Case 1
            
                If UserList(UserIndex).Counters.Pena <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If
                
                
                If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = CARCEL Then
                    Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                End If
            
            
                Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                Call WriteConsoleMsg(UserIndex, "Has viajado por el mundo.", FontTypeNames.FONTTYPE_WARNING)
                Call QuitarUserInvItem(UserIndex, slot, 1)
                Call UpdateUserInv(False, UserIndex, slot)
                
            Case 2
                If DonadorCheck(UserList(UserIndex).Cuenta) = 0 Then
                    Call DonadorTiempo(UserList(UserIndex).Cuenta, CLng(obj.CuantoAumento))
                    Call WriteConsoleMsg(UserIndex, "Donaci�n> Se han agregado " & obj.CuantoAumento & " dias de donador a tu cuenta. Relogea tu personaje para empezar a disfrutar la experiencia.", FontTypeNames.FONTTYPE_WARNING)
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    Call UpdateUserInv(False, UserIndex, slot)
                Else
                    Call DonadorTiempo(UserList(UserIndex).Cuenta, CLng(obj.CuantoAumento))
                    Call WriteConsoleMsg(UserIndex, "�Se han a�adido " & CLng(obj.CuantoAumento) & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
                    UserList(UserIndex).donador.activo = 1
                    Call QuitarUserInvItem(UserIndex, slot, 1)
                    Call UpdateUserInv(False, UserIndex, slot)
                    'Call WriteConsoleMsg(UserIndex, "Donaci�n> Debes esperar a que finalice el periodo existente para renovar tu suscripci�n.", FontTypeNames.FONTTYPE_INFOIAO)
                End If
            Case 3
                Call AgregarCreditosDonador(UserList(UserIndex).Cuenta, CLng(obj.CuantoAumento))
                Call WriteConsoleMsg(UserIndex, "Donaci�n> Tu credito ahora es de " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                Call QuitarUserInvItem(UserIndex, slot, 1)
                Call UpdateUserInv(False, UserIndex, slot)
        End Select
    
        
     
    Case eOBJType.otpasajes
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.TargetNpcTipo <> Pirata Then
           Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
            Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
            'Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Pos.Map <> obj.DesdeMap Then
          Rem  Call WriteConsoleMsg(UserIndex, "El pasaje no lo compraste aqu�! Largate!", FontTypeNames.FONTTYPE_INFO)
            Call WriteChatOverHead(UserIndex, "El pasaje no lo compraste aqu�! Largate!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        
        If Not MapaValido(obj.HastaMap) Then
           Rem Call WriteConsoleMsg(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
            Call WriteChatOverHead(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        If obj.NecesitaNave > 0 Then
            If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
                Call WriteChatOverHead(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
            
        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
        Call WriteConsoleMsg(UserIndex, "Has viajado por varios d�as, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
        UserList(UserIndex).Stats.MinAGU = 0
        UserList(UserIndex).Stats.MinHam = 0
        UserList(UserIndex).flags.Sed = 1
        UserList(UserIndex).flags.Hambre = 1
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call QuitarUserInvItem(UserIndex, slot, 1)
        Call UpdateUserInv(False, UserIndex, slot)
        
        
        
    Case eOBJType.otRunas
    
        If UserList(UserIndex).Counters.Pena <> 0 Then
            Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = CARCEL Then
            Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.BattleModo = 1 Then
            Call WriteConsoleMsg(UserIndex, "No podes usarlo aqu�.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        
        If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 And UserList(UserIndex).flags.Muerto = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solo podes usar tu runa en zonas seguras.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).accion.AccionPendiente Then
            Exit Sub
        End If
        
        Select Case ObjData(ObjIndex).TipoRuna
        
        Case 1, 2

            If UserList(UserIndex).donador.activo = 0 Then ' Donador no espera tiempo
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 400, False))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 350, Accion_Barra.Runa))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 50, False))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 100, Accion_Barra.Runa))
            End If
            UserList(UserIndex).accion.Particula = ParticulasIndex.Runa
            UserList(UserIndex).accion.AccionPendiente = True
            UserList(UserIndex).accion.TipoAccion = Accion_Barra.Runa
            UserList(UserIndex).accion.RunaObj = ObjIndex
            UserList(UserIndex).accion.ObjSlot = slot
            
        Case 3
        
         Dim parejaindex As Integer


        If Not UserList(UserIndex).flags.BattleModo Then
                
            'If UserList(UserIndex).donador.activo = 1 Then
                If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                    If UserList(UserIndex).flags.Casado = 1 Then
                        parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                        
                            If parejaindex > 0 Then
                                If UserList(parejaindex).flags.BattleModo = 0 Then
                                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 600, False))
                                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.GoToPareja))
                                    UserList(UserIndex).accion.AccionPendiente = True
                                    UserList(UserIndex).accion.Particula = ParticulasIndex.Runa
                                    UserList(UserIndex).accion.TipoAccion = Accion_Barra.GoToPareja
                                Else
                                    Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No pod�s teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)
                                End If
                                
                            Else
                                Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)
                            End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)
                End If
                
           ' Else
              '  Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)
           ' End If
        Else
            Call WriteConsoleMsg(UserIndex, "No pod�s usar esta opci�n en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
        
        End If

    
        End Select
        
         Case eOBJType.otmapa
            Call WriteShowFrmMapa(UserIndex)
            
        
End Select

Exit Sub


hErr:
    LogError "Error en useinvitem Usuario: " & UserList(UserIndex).name & " item:" & obj.name & " index: " & UserList(UserIndex).Invent.Object(slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub
Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)

Call WriteAlquimistaObjects(UserIndex)

End Sub
Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)

Call WriteSastreObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next


If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub
If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub


Call TirarTodosLosItems(UserIndex)



End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And _
            (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And _
            ObjData(Index).OBJType <> eOBJType.otLlaves And _
            ObjData(Index).OBJType <> eOBJType.otBarcos And _
            ObjData(Index).OBJType <> eOBJType.otMonturas And _
            ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And _
            ObjData(Index).donador = 0 And Not ObjData(Index).Instransferible = 1
             


End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As obj
    Dim ItemIndex As Integer
    
    

    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
    
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
                    

             If ItemSeCae(ItemIndex) Then
                NuevaPos.x = 0
                NuevaPos.Y = 0
 
                If ItemIndex = ORO_MINA And UserList(UserIndex).flags.CarroMineria = 1 Or ItemIndex = PLATA_MINA And UserList(UserIndex).flags.CarroMineria = 1 Or ItemIndex = HIERRO_MINA And UserList(UserIndex).flags.CarroMineria = 1 Then
                    MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount * 0.3
                    MiObj.ObjIndex = ItemIndex
                    
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MiObj.Amount, NuevaPos.Map, NuevaPos.x, NuevaPos.Y)
                    End If
                
                Else
            
                    MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
                
                    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.x, NuevaPos.Y)
                    End If
                End If
              
             End If
        End If
      
    
    Next i

End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 6 Then Exit Sub

For i = 1 To UserList(UserIndex).CurrentInventorySlots
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.x = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.x, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.x, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
