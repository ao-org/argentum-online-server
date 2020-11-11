Attribute VB_Name = "modSistemaComercio"
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
Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal slot As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo Comercio_Err
        

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 27/07/08 (MarKoxX) |
        '27/07/08 (MarKoxX) - New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
        '06/13/08 (NicoNZ)
        '24/01/2020: WyroX = Reduzco la cantidad de paquetes que se envian, actualizo solo los slots necesarios y solo el oro, no todos los stats.
        '*************************************************
        Dim Precio       As Long

        Dim Objeto       As obj

        Dim objquedo     As obj

        Dim precioenvio  As Single

        Dim DestruirItem As Boolean
    
100     DestruirItem = False
    
        Dim NpcSlot As Integer
    
102     If Cantidad < 1 Or slot < 1 Then Exit Sub
    
104     If Modo = eModoComercio.Compra Then
106         If slot > UserList(UserIndex).CurrentInventorySlots Then
                Exit Sub
108         ElseIf Cantidad > MAX_INVENTORY_OBJS Then
110             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
112             Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
114             UserList(UserIndex).flags.Ban = 1
116             Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            
118             Call CloseSocket(UserIndex)
                Exit Sub
120         ElseIf Not Npclist(NpcIndex).Invent.Object(slot).Amount > 0 Then
                Exit Sub

            End If
        
122         If Cantidad > Npclist(NpcIndex).Invent.Object(slot).Amount Then Cantidad = Npclist(NpcIndex).Invent.Object(slot).Amount
        
124         Objeto.Amount = Cantidad
126         Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(slot).ObjIndex
        
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
128         Precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(slot).ObjIndex).Valor / Descuento(UserIndex) * Cantidad) + 0.5)
        
130         If UserList(UserIndex).Stats.GLD < Precio Then
132             Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
135         If Not MeterItemEnInventario(UserIndex, Objeto) Then Exit Sub
        
140         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
            
            Call WriteUpdateGold(UserIndex)
    
            Call QuitarNpcInvItem(NpcIndex, slot, Cantidad)
            Call UpdateNpcInvToAll(False, NpcIndex, slot)            'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
            
            'Es un Objeto que tenemos que loguear?
            'If ObjData(Objeto.ObjIndex).Log = 1 Then
            '   Call LogDesarrollo(UserList(UserIndex).name & " compr� del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            'ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            '   If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
            ''     Call LogDesarrollo(UserList(UserIndex).name & " compr� del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            '  End If
            'End If
        
            'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
144         If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
146             Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & slot, Objeto.ObjIndex & "-0")
148             Call logVentaCasa(UserList(UserIndex).name & " compro " & ObjData(Objeto.ObjIndex).name)

            End If
         
            Rem    NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
                
150         objquedo.Amount = Npclist(NpcIndex).Invent.Object(CByte(slot)).Amount
152         objquedo.ObjIndex = Npclist(NpcIndex).Invent.Object(CByte(slot)).ObjIndex
154         precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
    
156         Call WriteChangeNPCInventorySlot(UserIndex, CByte(slot), objquedo, precioenvio)
            
            Rem    precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
    
            Rem Call WriteChangeNPCInventorySlot(UserIndex, NpcSlot, objquedo, precioenvio)
        
158     ElseIf Modo = eModoComercio.Venta Then
        
160         If Cantidad > UserList(UserIndex).Invent.Object(slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(slot).Amount
        
162         Objeto.Amount = Cantidad
164         Objeto.ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex

166         If Objeto.ObjIndex = 0 Then
                Exit Sub
                
168         ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
170             Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                
176         ElseIf ObjData(Objeto.ObjIndex).Destruye = 1 Then
178             Call WriteConsoleMsg(UserIndex, "Lo siento, no puedo comprarte ese item.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

184         ElseIf UserList(UserIndex).flags.BattleModo = 1 Then
186             Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio items robados.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
192         ElseIf Not MeterItemEnInventarioDeNpc(UserList(UserIndex).flags.TargetNPC, Objeto) Then
194             DestruirItem = True
          
196         ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
198             Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
204
            ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then

206             If Npclist(NpcIndex).name <> "SR" Then
208                 Call WriteConsoleMsg(UserIndex, "Las armaduras de la Armada solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

214         ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

216             If Npclist(NpcIndex).name <> "SC" Then
218                 Call WriteConsoleMsg(UserIndex, "Las armaduras de la Legi�n solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

224         ElseIf UserList(UserIndex).Invent.Object(slot).Amount < 0 Or Cantidad = 0 Then
                Exit Sub
                
226         ElseIf slot < LBound(UserList(UserIndex).Invent.Object()) Or slot > UBound(UserList(UserIndex).Invent.Object()) Then
                Exit Sub
                
230         ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
232             Call WriteConsoleMsg(UserIndex, "No pod�s vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
238         Call QuitarUserInvItem(UserIndex, slot, Cantidad)
            
            Call UpdateUserInv(False, UserIndex, slot)
            
            'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
240         Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        
242         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
244         If UserList(UserIndex).Stats.GLD > MAXORO Then UserList(UserIndex).Stats.GLD = MAXORO
            
            Call WriteUpdateGold(UserIndex)
            
246         If DestruirItem Then

248             Call UpdateUserInv(False, UserIndex, slot)

250             Call WriteUpdateUserStats(UserIndex)

256             Call SubirSkill(UserIndex, eSkill.Comerciar)
            
                Exit Sub

            End If
        
258         NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
        
260         If NpcSlot <= UserList(UserIndex).CurrentInventorySlots Then 'Slot valido
                
                'Mete el obj en el slot
262             Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
264             Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount

266             If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
268                 Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
                End If
                
                Call UpdateNpcInvToAll(False, NpcIndex, NpcSlot)
                
            End If
        
            'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            ' If ObjData(Objeto.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " vendi� al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            ' ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            '     If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
            '         Call LogDesarrollo(UserList(UserIndex).name & " vendi� al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            '     End If
            ' End If
            
270         objquedo.Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount
    
272         objquedo.ObjIndex = Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex
    
274         precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
  
276         Call WriteChangeNPCInventorySlot(UserIndex, NpcSlot, objquedo, precioenvio)

        End If

286     Call SubirSkill(UserIndex, eSkill.Comerciar)

        
        Exit Sub

Comercio_Err:

        Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.Comercio", Erl)
        Resume Next
        
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        
        On Error GoTo IniciarComercioNPC_Err
        

100     Call UpdateNpcInv(True, UserIndex, UserList(UserIndex).flags.TargetNPC, 0)

102     If Npclist(UserList(UserIndex).flags.TargetNPC).SoundOpen <> 0 Then
104         Call WritePlayWave(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)
        End If

106     UserList(UserIndex).flags.Comerciando = True

108     Call WriteCommerceInit(UserIndex)

        Exit Sub

IniciarComercioNPC_Err:

        Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.IniciarComercioNPC", Erl)
        Resume Next
        
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        
        On Error GoTo SlotEnNPCInv_Err
        
100     SlotEnNPCInv = 1

102     Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
104         SlotEnNPCInv = SlotEnNPCInv + 1

106         If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
        Loop
    
108     If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
110         SlotEnNPCInv = 1
        
112         Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
114             SlotEnNPCInv = SlotEnNPCInv + 1

116             If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
            Loop
        
118         If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    
        End If
    
        
        Exit Function

SlotEnNPCInv_Err:
        Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.SlotEnNPCInv", Erl)
        Resume Next
        
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 2/8/06
        '*************************************************
        
        On Error GoTo Descuento_Err
        
100     Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100

        
        Exit Function

Descuento_Err:
        Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.Descuento", Erl)
        Resume Next
        
End Function

''
' Update the inventory of the Npc to the user
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC

Private Sub UpdateNpcInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo EnviarNpcInv_Err

        Dim obj As obj
        Dim LoopC As Long
        Dim Desc As Single
        Dim val As Single
        
        Desc = Descuento(UserIndex)
        
        'Actualiza un solo slot
        If Not UpdateAll Then
        
            With Npclist(NpcIndex).Invent.Object(slot)
                obj.ObjIndex = .ObjIndex
                obj.Amount = .Amount
                
                If .ObjIndex > 0 Then
                    val = (ObjData(.ObjIndex).Valor) / Desc
                End If
                
                Call WriteChangeNPCInventorySlot(UserIndex, slot, obj, val)
                
            End With
            
        Else
        
            'Actualiza todos los slots
            For LoopC = 1 To MAX_INVENTORY_SLOTS
            
                With Npclist(NpcIndex).Invent.Object(LoopC)
                
                    obj.ObjIndex = .ObjIndex
                    obj.Amount = .Amount
    
                    If .ObjIndex > 0 Then
                        val = (ObjData(.ObjIndex).Valor) / Desc
                    End If
    
                    Call WriteChangeNPCInventorySlot(UserIndex, LoopC, obj, val)
                    
                End With
                
            Next LoopC
            
        End If

        Exit Sub

EnviarNpcInv_Err:
        Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.UpdateNpcInv", Erl)
        Resume Next
        
End Sub

''
' Update the inventory of the Npc to all users trading with him
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC
' @param slot The slot to update

Public Sub UpdateNpcInvToAll(ByVal UpdateAll As Boolean, ByVal NpcIndex As Integer, ByVal slot As Byte)
'***************************************************
On Error GoTo Errhandler:

    Dim LoopC As Long

    ' Recorremos todos los usuarios
    For LoopC = 1 To LastUser
    
        With UserList(LoopC)
        
            ' Si esta comerciando
            If .flags.Comerciando Then
            
                ' Si el ultimo NPC que cliqueo es el que hay que actualizar
                If .flags.TargetNPC = NpcIndex Then
                    ' Actualizamos el inventario del NPC
                    Call UpdateNpcInv(UpdateAll, LoopC, NpcIndex, slot)
                End If
                
            End If
            
        End With
        
    Next
    
    Exit Sub
    
Errhandler:
    
    Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.UpdateNpcInvToAll")
    Resume Next
    
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param valor  El valor de compra de objeto

Public Function SalePrice(ByVal ObjIndex As Integer) As Single
        
        On Error GoTo SalePrice_Err
        

        '*************************************************
        'Author: Nicol�s (NicoNZ)
        '
        '*************************************************
100     If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
102     If ItemNewbie(ObjIndex) Then Exit Function
    
104     SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA

        
        Exit Function

SalePrice_Err:
        Call RegistrarError(Err.Number, Err.description, "modSistemaComercio.SalePrice", Erl)
        Resume Next
        
End Function

