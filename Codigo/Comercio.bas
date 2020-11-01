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
'*************************************************
'Author: Nacho (Integer)
'Last modified: 27/07/08 (MarKoxX) | New changes in the way of trading (now when you buy it rounds to ceil and when you sell it rounds to floor)
'  - 06/13/08 (NicoNZ)
'*************************************************
    Dim Precio As Long
    Dim Objeto As obj
    Dim objquedo As obj
    Dim precioenvio As Single
    Dim DestruirItem As Boolean
    
    DestruirItem = False
    
 Dim NpcSlot As Integer
    
    If Cantidad < 1 Or slot < 1 Then Exit Sub

    
    If Modo = eModoComercio.Compra Then
        If slot > UserList(UserIndex).CurrentInventorySlots Then
            Exit Sub
        ElseIf Cantidad > MAX_INVENTORY_OBJS Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
            Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
            UserList(UserIndex).flags.Ban = 1
            Call WriteErrorMsg(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        ElseIf Not Npclist(NpcIndex).Invent.Object(slot).Amount > 0 Then
            Exit Sub
        End If
        
        If Cantidad > Npclist(NpcIndex).Invent.Object(slot).Amount Then Cantidad = Npclist(UserList(UserIndex).flags.TargetNPC).Invent.Object(slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = Npclist(NpcIndex).Invent.Object(slot).ObjIndex
        
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        Precio = CLng((ObjData(Npclist(NpcIndex).Invent.Object(slot).ObjIndex).Valor / Descuento(UserIndex) * Cantidad) + 0.5)

        
        If UserList(UserIndex).Stats.GLD < Precio Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If MeterItemEnInventario(UserIndex, Objeto) = False Then
            'Call WriteConsoleMsg(UserIndex, "No pod�s cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
        
        Call QuitarNpcInvItem(UserList(UserIndex).flags.TargetNPC, CByte(slot), Cantidad)
        
        'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
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
        If ObjData(Objeto.ObjIndex).OBJType = otLlaves Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & Npclist(NpcIndex).Numero, "obj" & slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(UserIndex).name & " compro " & ObjData(Objeto.ObjIndex).name)
        End If
              
         
              
         
    Rem    NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
                
    objquedo.Amount = Npclist(NpcIndex).Invent.Object(CByte(slot)).Amount
    objquedo.ObjIndex = Npclist(NpcIndex).Invent.Object(CByte(slot)).ObjIndex
    precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
    
      Call WriteChangeNPCInventorySlot(UserIndex, CByte(slot), objquedo, precioenvio)
            
        Rem    precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
    
     Rem Call WriteChangeNPCInventorySlot(UserIndex, NpcSlot, objquedo, precioenvio)
        
    ElseIf Modo = eModoComercio.Venta Then
    
        
        
        If Cantidad > UserList(UserIndex).Invent.Object(slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(slot).Amount
        
        Objeto.Amount = Cantidad
        Objeto.ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
        If Objeto.ObjIndex = 0 Then
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio objetos para newbies.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Destruye = 1 Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no puedo comprarte ese item.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf UserList(UserIndex).flags.BattleModo = 1 Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio items robados.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        
            
        ElseIf Not MeterItemEnInventarioDeNpc(UserList(UserIndex).flags.TargetNPC, Objeto) Then
            DestruirItem = True
            'Call WriteConsoleMsg(UserIndex, "Lo siento, no puedo cargar mas items...", FontTypeNames.FONTTYPE_INFO)
            'Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
           ' Call WriteTradeOK(UserIndex)
            'Exit Sub
          
        ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
            Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then
            If Npclist(NpcIndex).name <> "SR" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras de la Armada solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then
            If Npclist(NpcIndex).name <> "SC" Then
                Call WriteConsoleMsg(UserIndex, "Las armaduras de la Legi�n solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
                Call WriteTradeOK(UserIndex)
                Exit Sub
            End If
        ElseIf UserList(UserIndex).Invent.Object(slot).Amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf slot < LBound(UserList(UserIndex).Invent.Object()) Or slot > UBound(UserList(UserIndex).Invent.Object()) Then
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        ElseIf UserList(UserIndex).flags.Privilegios And PlayerType.Consejero Then
            Call WriteConsoleMsg(UserIndex, "No pod�s vender items.", FontTypeNames.FONTTYPE_WARNING)
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
            Exit Sub
        End If
        
        Call QuitarUserInvItem(UserIndex, slot, Cantidad)
        
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
        If UserList(UserIndex).Stats.GLD > MAXORO Then _
            UserList(UserIndex).Stats.GLD = MAXORO
        
        
        If DestruirItem Then
            Call UpdateUserInv(False, UserIndex, slot)
            Call WriteUpdateUserStats(UserIndex)
            
            Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Call WriteTradeOK(UserIndex)
                
            Call SubirSkill(UserIndex, eSkill.Comerciar)
            
            Exit Sub
        End If
        
        NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
        
        If NpcSlot <= UserList(UserIndex).CurrentInventorySlots Then 'Slot valido
            'Mete el obj en el slot
            Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
            Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount
            If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
                Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
            End If
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
          objquedo.Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount
    
    objquedo.ObjIndex = Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex
    
    precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
  
      Call WriteChangeNPCInventorySlot(UserIndex, NpcSlot, objquedo, precioenvio)
    End If
               


    
    Call UpdateUserInv(False, UserIndex, slot)
    Call WriteUpdateUserStats(UserIndex)
    
    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    Call WriteTradeOK(UserIndex)
        
    Call SubirSkill(UserIndex, eSkill.Comerciar)
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************

    Call EnviarNpcInv(UserIndex, UserList(UserIndex).flags.TargetNPC)
    If Npclist(UserList(UserIndex).flags.TargetNPC).SoundOpen <> 0 Then
        Call WritePlayWave(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).SoundOpen, NO_3D_SOUND, NO_3D_SOUND)
    End If
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex)
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    SlotEnNPCInv = 1
    Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = Objeto _
      And Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        SlotEnNPCInv = SlotEnNPCInv + 1
        If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
        
    Loop
    
    If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then
    
        SlotEnNPCInv = 1
        
        Do Until Npclist(NpcIndex).Invent.Object(SlotEnNPCInv).ObjIndex = 0
        
            SlotEnNPCInv = SlotEnNPCInv + 1
            If SlotEnNPCInv > MAX_INVENTORY_SLOTS Then Exit Do
            
        Loop
        
        If SlotEnNPCInv <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    
    End If
    
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
'*************************************************
'Author: Nacho (Integer)
'Last modified: 2/8/06
'*************************************************
    Descuento = 1 + UserList(UserIndex).Stats.UserSkills(eSkill.Comerciar) / 100
End Function

''
' Send the inventory of the Npc to the user
'
' @param userIndex The index of the User
' @param npcIndex The index of the NPC

Private Sub EnviarNpcInv(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last Modified: 06/14/08
'Last Modified By: Nicol�s Ezequiel Bouhid (NicoNZ)
'*************************************************
    Dim slot As Byte
    Dim val As Single
    
    For slot = 1 To UserList(UserIndex).CurrentInventorySlots
        If Npclist(NpcIndex).Invent.Object(slot).ObjIndex > 0 Then
            Dim thisObj As obj
            thisObj.ObjIndex = Npclist(NpcIndex).Invent.Object(slot).ObjIndex
            thisObj.Amount = Npclist(NpcIndex).Invent.Object(slot).Amount
            val = CLng((ObjData(Npclist(NpcIndex).Invent.Object(slot).ObjIndex).Valor) / Descuento(UserIndex))
        
            Call WriteChangeNPCInventorySlot(UserIndex, slot, thisObj, val)
        Else
            Dim DummyObj As obj
            Call WriteChangeNPCInventorySlot(UserIndex, slot, DummyObj, 0)
        End If
    Next slot
    
    
End Sub

''
' Devuelve el valor de venta del objeto
'
' @param valor  El valor de compra de objeto

Public Function SalePrice(ByVal ObjIndex As Integer) As Single
'*************************************************
'Author: Nicol�s (NicoNZ)
'
'*************************************************
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    
    SalePrice = ObjData(ObjIndex).Valor / REDUCTOR_PRECIOVENTA
End Function


