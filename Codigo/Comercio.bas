Attribute VB_Name = "modSistemaComercio"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
''*****************************************************
'Sistema de Comercio para Argentum Online
'Programado por Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Enum eModoComercio

    Compra = 1
    Venta = 2

End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

''
' Makes a trade. (Buy or Sell)
'
' @param Modo The trade type (sell or buy)
' @param UserIndex Specifies the index of the user
' @param NpcIndex specifies the index of the npc
' @param Slot Specifies which slot are you trying to sell / buy
' @param Cantidad Specifies how many items in that slot are you trying to sell / buy
Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo Comercio_Err
        

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 27/07/08 (MarKoxX) |
        '27/07/08 (MarKoxX) - New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
        '06/13/08 (NicoNZ)
        '24/01/2020: WyroX = Reduzco la cantidad de paquetes que se envian, actualizo solo los slots necesarios y solo el oro, no todos los stats.
        '*************************************************
        Dim Precio       As Long
        Dim Objeto       As t_Obj
        Dim objquedo     As t_Obj
        Dim precioenvio  As Single
        Dim NpcSlot As Integer
        
        Dim Objeto_A_Comprar As t_UserOBJ
        
100     If Cantidad < 1 Or Slot < 1 Then Exit Sub
    
102     If Modo = eModoComercio.Compra Then

104         If Slot > MAX_INVENTORY_SLOTS Then
                Exit Sub
                
106         ElseIf Cantidad > MAX_INVENTORY_OBJS Then
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " ha sido baneado por el sistema anti-cheats.", e_FontTypeNames.FONTTYPE_FIGHT))
110             Call Ban(UserList(UserIndex).Name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
112             UserList(UserIndex).flags.Ban = 1
114             Call WriteShowMessageBox(UserIndex, "Has sido baneado por el Sistema AntiCheat.")

116             Call CloseSocket(UserIndex)
                Exit Sub
                
118         ElseIf Not NpcList(NpcIndex).Invent.Object(Slot).amount > 0 Then
                Exit Sub

            End If
            
            Objeto_A_Comprar = NpcList(NpcIndex).Invent.Object(Slot)
            
            If Objeto_A_Comprar.ObjIndex = 0 Then Exit Sub
            
            If ObjData(Objeto_A_Comprar.ObjIndex).Crucial = 0 Then
                If Cantidad > Objeto_A_Comprar.amount Then
                    Cantidad = Objeto_A_Comprar.amount
                End If
            ElseIf ObjData(Objeto_A_Comprar.ObjIndex).Crucial = 1 Then
                'si es un item que vende el NPC le dejo comprar todo lo que quiera
                If Not NpcSellsItem(NpcList(NpcIndex).Numero, Objeto_A_Comprar.ObjIndex) Then
                    If Cantidad > Objeto_A_Comprar.amount Then
                        Cantidad = Objeto_A_Comprar.amount
                    End If
                End If
            End If
            
            

            'NpcSellsItem
            
            
122         Objeto.amount = Cantidad
124         Objeto.ObjIndex = Objeto_A_Comprar.ObjIndex
        
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
126         Precio = Ceil(ObjData(Objeto_A_Comprar.ObjIndex).Valor / Descuento(UserIndex) * Cantidad)
        
128         If UserList(UserIndex).Stats.GLD < Precio Then
130             Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
132         If Not MeterItemEnInventario(UserIndex, Objeto) Then Exit Sub
        
134         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
            
136         Call WriteUpdateGold(UserIndex)
    
138         Call QuitarNpcInvItem(NpcIndex, Slot, Cantidad)
140         Call UpdateNpcInvToAll(False, NpcIndex, Slot)

            'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
142         If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
144             Call WriteVar(DatPath & "NPCs.dat", "NPC" & NpcList(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
146             Call logVentaCasa(UserList(UserIndex).Name & " compro " & ObjData(Objeto.ObjIndex).Name)

            End If
        
156     ElseIf Modo = eModoComercio.Venta Then

            If Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
158         If Cantidad > UserList(UserIndex).Invent.Object(Slot).amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).amount
        
160         Objeto.amount = Cantidad
162         Objeto.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex

164         If Objeto.ObjIndex = 0 Then
                Exit Sub
                
166         ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
168             Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio objetos para newbies.", e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub
                
170         ElseIf ObjData(Objeto.ObjIndex).Destruye = 1 Then
172             Call WriteConsoleMsg(UserIndex, "Lo siento, no puedo comprarte ese item.", e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            
174         ElseIf ObjData(Objeto.ObjIndex).Instransferible = 1 Then
176             Call WriteConsoleMsg(UserIndex, "Lo siento, no puedo comprarte ese item.", e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub
          
178         ElseIf (NpcList(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And NpcList(NpcIndex).TipoItems <> e_OBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
180             Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub

182         ElseIf UserList(UserIndex).Invent.Object(Slot).amount < 0 Or Cantidad = 0 Then
                Exit Sub
                
184         ElseIf Slot < LBound(UserList(UserIndex).Invent.Object()) Or Slot > UBound(UserList(UserIndex).Invent.Object()) Then
                Exit Sub
                
186         ElseIf UserList(UserIndex).flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
188             Call WriteConsoleMsg(UserIndex, "No podés vender items.", e_FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
190         Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
            
192         Call UpdateUserInv(False, UserIndex, Slot)
            
            'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
194         Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        
196         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
198         If UserList(UserIndex).Stats.GLD > MAXORO Then UserList(UserIndex).Stats.GLD = MAXORO
            
200         Call WriteUpdateGold(UserIndex)
        
202         NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.amount)
        
204         If NpcSlot > 0 And NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                
                ' Saque este incremento de SlotEnNPCInv porque me parece mejor manejarlo junto con el resto de las asignaciones
206             If NpcList(NpcIndex).Invent.Object(NpcSlot).ObjIndex = 0 Then
208                 NpcList(NpcIndex).Invent.NroItems = NpcList(NpcIndex).Invent.NroItems + 1
                End If
                
                'Mete el obj en el slot
210             NpcList(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
212             NpcList(NpcIndex).Invent.Object(NpcSlot).amount = NpcList(NpcIndex).Invent.Object(NpcSlot).amount + Objeto.amount

214             If NpcList(NpcIndex).Invent.Object(NpcSlot).amount > MAX_INVENTORY_OBJS Then
216                 NpcList(NpcIndex).Invent.Object(NpcSlot).amount = MAX_INVENTORY_OBJS
                End If
                
218             Call UpdateNpcInvToAll(False, NpcIndex, NpcSlot)
                
            End If

        End If

228     Call SubirSkill(UserIndex, e_Skill.Comerciar)

        Exit Sub

Comercio_Err:

230     Call TraceError(Err.Number, Err.Description, "modSistemaComercio.Comercio", Erl)

        
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        
        On Error GoTo IniciarComercioNPC_Err
        
        If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
            Call WriteConsoleMsg(UserIndex, "El comerciante no esta disponible.", e_FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
100     Call UpdateNpcInv(True, UserIndex, UserList(UserIndex).flags.TargetNPC.ArrayIndex, 0)

102     If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundOpen <> 0 Then
104         Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)
        End If

106     UserList(UserIndex).flags.Comerciando = True

108     Call WriteCommerceInit(UserIndex)

        Exit Sub

IniciarComercioNPC_Err:

110     Call TraceError(Err.Number, Err.Description, "modSistemaComercio.IniciarComercioNPC", Erl)

        
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
        '*************************************************
        'Devuelve el slot en el cual se debe agregar el nuevo objeto, o 0 si no se debe asignar en ningun lado
        '*************************************************
        
        On Error GoTo SlotEnNPCInv_Err
                       
100     With NpcList(NpcIndex).Invent
        
            Dim Slot As Byte
            Dim matchingSlots As New Collection
            Dim firstEmptySpace As Integer
            
            ' Recorro el inventario buscando el objeto a agregar y espacios vacios
102         firstEmptySpace = 0
104         For Slot = 1 To MAX_INVENTORY_SLOTS
106             If .Object(Slot).ObjIndex = Objeto Then
108                 matchingSlots.Add (Slot)
110             ElseIf .Object(Slot).ObjIndex = 0 And firstEmptySpace = 0 Then
112                 firstEmptySpace = Slot
                End If
114         Next Slot
            
            ' Recorro los slots donde hay objetos que matcheen con el objeto a agregar y si alguno tiene espacio, lo agrego ahi. Si no, se descarta
116         If matchingSlots.Count <> 0 Then
                Dim i As Variant
118             For Each i In matchingSlots
120                 If .Object(i).amount < MAX_INVENTORY_OBJS Then
122                     SlotEnNPCInv = i
                        Exit Function
                    End If
124             Next i
126             SlotEnNPCInv = 0
                Exit Function
            End If
            
128         SlotEnNPCInv = firstEmptySpace
            Exit Function
        End With
        
SlotEnNPCInv_Err:
130     Call TraceError(Err.Number, Err.Description, "modSistemaComercio.SlotEnNPCInv", Erl)

        
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        
        On Error GoTo Descuento_Err
        
100     Descuento = 1 + UserList(UserIndex).Stats.UserSkills(e_Skill.Comerciar) / 100

        
        Exit Function

Descuento_Err:
102     Call TraceError(Err.Number, Err.Description, "modSistemaComercio.Descuento", Erl)

        
End Function

''
' Update the inventory of the Npc to the user
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC

Private Sub UpdateNpcInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Byte)
        
            On Error GoTo EnviarNpcInv_Err

            Dim obj As t_Obj
            Dim LoopC As Long
            Dim Desc As Single
            Dim val As Single
        
100         Desc = Descuento(UserIndex)
        
            'Actualiza un solo slot
102         If Not UpdateAll Then
        
104             With NpcList(NpcIndex).Invent.Object(Slot)
106                 obj.ObjIndex = .ObjIndex
108                 obj.amount = .amount
                
110                 If .ObjIndex > 0 Then
112                     val = Ceil(ObjData(.ObjIndex).Valor / Desc)
                    End If
                
114                 Call WriteChangeNPCInventorySlot(UserIndex, Slot, obj, val)
                
                End With
            
            Else
        
                'Actualiza todos los slots
116             For LoopC = 1 To MAX_INVENTORY_SLOTS
            
118                 With NpcList(NpcIndex).Invent.Object(LoopC)
                
120                     obj.ObjIndex = .ObjIndex
122                     obj.amount = .amount
    
124                     If .ObjIndex > 0 Then
126                         val = Ceil(ObjData(.ObjIndex).Valor / Desc)
                        End If
    
128                     Call WriteChangeNPCInventorySlot(UserIndex, LoopC, obj, val)
                    
                    End With
                
130             Next LoopC
            
            End If

            Exit Sub

EnviarNpcInv_Err:
132         Call TraceError(Err.Number, Err.Description, "modSistemaComercio.UpdateNpcInv", Erl)

        
End Sub

''
' Update the inventory of the Npc to all users trading with him
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC
' @param slot The slot to update

Public Sub UpdateNpcInvToAll(ByVal UpdateAll As Boolean, ByVal NpcIndex As Integer, ByVal Slot As Byte)
    '***************************************************
    On Error GoTo ErrHandler:

        Dim LoopC As Long

        ' Recorremos todos los usuarios
100     For LoopC = 1 To LastUser
    
102         With UserList(LoopC)
        
                ' Si esta comerciando
104             If .flags.Comerciando Then
            
                    ' Si el ultimo NPC que cliqueo es el que hay que actualizar
106                 If .flags.TargetNPC.ArrayIndex = NpcIndex Then
                        ' Actualizamos el inventario del NPC
108                     Call UpdateNpcInv(UpdateAll, LoopC, NpcIndex, Slot)
                    End If
                
                End If
            
            End With
        
        Next
    
        Exit Sub
    
ErrHandler:
    
110     Call TraceError(Err.Number, Err.Description, "modSistemaComercio.UpdateNpcInvToAll")

    
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param valor  El valor de compra de objeto

Public Function SalePrice(ByVal ObjIndex As Integer) As Single
        
        On Error GoTo SalePrice_Err
        

        '*************************************************
        'Author: Nicolás (NicoNZ)
        '
        '*************************************************
100     If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
102     If ItemNewbie(ObjIndex) Then Exit Function
    
104     SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA

        
        Exit Function

SalePrice_Err:
106     Call TraceError(Err.Number, Err.Description, "modSistemaComercio.SalePrice", Erl)

        
End Function

