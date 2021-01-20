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
    
        Dim NpcSlot As Integer
    
102     If Cantidad < 1 Or slot < 1 Then Exit Sub
    
104     If Modo = eModoComercio.Compra Then
106         If slot > UserList(UserIndex).CurrentInventorySlots Then
                Exit Sub
108         ElseIf Cantidad > MAX_INVENTORY_OBJS Then
110             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " ha sido baneado por el sistema anti-cheats.", FontTypeNames.FONTTYPE_FIGHT))
112             Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
114             UserList(UserIndex).flags.Ban = 1
116             Call WriteShowMessageBox(UserIndex, "Has sido baneado por el Sistema AntiCheat.")
            
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
            
134         If Not MeterItemEnInventario(UserIndex, Objeto) Then Exit Sub
        
136         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
            
138         Call WriteUpdateGold(UserIndex)
    
140         Call QuitarNpcInvItem(NpcIndex, slot, Cantidad)
142         Call UpdateNpcInvToAll(False, NpcIndex, slot)            'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
            
            'Es un Objeto que tenemos que loguear?
            'If ObjData(Objeto.ObjIndex).Log = 1 Then
            '   Call LogDesarrollo(UserList(UserIndex).name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            'ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            '   If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
            ''     Call LogDesarrollo(UserList(UserIndex).name & " compró del NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
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
                
172         ElseIf ObjData(Objeto.ObjIndex).Destruye = 1 Then
174             Call WriteConsoleMsg(UserIndex, "Lo siento, no puedo comprarte ese item.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

176         ElseIf UserList(UserIndex).flags.BattleModo = 1 Then
178             Call WriteConsoleMsg(UserIndex, "Lo siento, no comercio items robados.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
          
184         ElseIf (Npclist(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera) Or Objeto.ObjIndex = iORO Then
186             Call WriteConsoleMsg(UserIndex, "Lo siento, no estoy interesado en este tipo de objetos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

188         ElseIf ObjData(Objeto.ObjIndex).Real = 1 Then

190             If Npclist(NpcIndex).name <> "SR" Then
192                 Call WriteConsoleMsg(UserIndex, "Las armaduras de la Armada solo pueden ser vendidas a los sastres reales.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

194         ElseIf ObjData(Objeto.ObjIndex).Caos = 1 Then

196             If Npclist(NpcIndex).name <> "SC" Then
198                 Call WriteConsoleMsg(UserIndex, "Las armaduras de la Legión solo pueden ser vendidas a los sastres del demonio.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

200         ElseIf UserList(UserIndex).Invent.Object(slot).Amount < 0 Or Cantidad = 0 Then
                Exit Sub
                
202         ElseIf slot < LBound(UserList(UserIndex).Invent.Object()) Or slot > UBound(UserList(UserIndex).Invent.Object()) Then
                Exit Sub
                
204         ElseIf UserList(UserIndex).flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
206             Call WriteConsoleMsg(UserIndex, "No podés vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
208         Call QuitarUserInvItem(UserIndex, slot, Cantidad)
            
210         Call UpdateUserInv(False, UserIndex, slot)
            
            'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
212         Precio = Fix(SalePrice(Objeto.ObjIndex) * Cantidad)
        
214         UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Precio
        
216         If UserList(UserIndex).Stats.GLD > MAXORO Then UserList(UserIndex).Stats.GLD = MAXORO
            
218         Call WriteUpdateGold(UserIndex)
        
228         NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
        
230         If NpcSlot > 0 And NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                
                ' Saque este incremento de SlotEnNPCInv porque me parece mejor manejarlo junto con el resto de las asignaciones
                If Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = 0 Then
                    Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
                End If
                
                'Mete el obj en el slot
232             Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
234             Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount + Objeto.Amount

236             If Npclist(NpcIndex).Invent.Object(NpcSlot).Amount > MAX_INVENTORY_OBJS Then
238                 Npclist(NpcIndex).Invent.Object(NpcSlot).Amount = MAX_INVENTORY_OBJS
                End If
                
240             Call UpdateNpcInvToAll(False, NpcIndex, NpcSlot)

242             objquedo.Amount = Npclist(NpcIndex).Invent.Object(NpcSlot).Amount
    
244             objquedo.ObjIndex = Npclist(NpcIndex).Invent.Object(NpcSlot).ObjIndex
    
246             precioenvio = CLng((ObjData(Objeto.ObjIndex).Valor / Descuento(UserIndex) * 1))
  
248             Call WriteChangeNPCInventorySlot(UserIndex, NpcSlot, objquedo, precioenvio)
                
            End If
        
            'Bien, ahora logueo de ser necesario. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            ' If ObjData(Objeto.ObjIndex).Log = 1 Then
            '    Call LogDesarrollo(UserList(UserIndex).name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            ' ElseIf Objeto.Amount = 1000 Then 'Es mucha cantidad?
            'Si no es de los prohibidos de loguear, lo logueamos.
            '     If ObjData(Objeto.ObjIndex).NoLog <> 1 Then
            '         Call LogDesarrollo(UserList(UserIndex).name & " vendió al NPC " & Objeto.Amount & " " & ObjData(Objeto.ObjIndex).name)
            '     End If
            ' End If

        End If

250     Call SubirSkill(UserIndex, eSkill.Comerciar)

        
        Exit Sub

Comercio_Err:

252     Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.Comercio", Erl)
254     Resume Next
        
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

110     Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.IniciarComercioNPC", Erl)
112     Resume Next
        
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
        '*************************************************
        'Devuelve el slot en el cual se debe agregar el nuevo objeto, o 0 si no se debe asignar en ningun lado
        '*************************************************
        
        On Error GoTo SlotEnNPCInv_Err
               
        Dim slot As Byte
        Dim matchingSlots As New Collection
        Dim firstEmptySpace As Integer
        
        ' Recorro el inventario buscando el objeto a agregar y espacios vacios
        firstEmptySpace = 0
        For slot = 1 To MAX_INVENTORY_SLOTS
            If Npclist(NpcIndex).Invent.Object(slot).ObjIndex = Objeto Then
                matchingSlots.Add (slot)
            ElseIf Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0 And firstEmptySpace = 0 Then
                firstEmptySpace = slot
            End If
        Next slot
        
        ' Recorro los slots donde hay objetos que matcheen con el objeto a agregar y si alguno tiene espacio, lo agrego ahi. Si no, se descarta
        If matchingSlots.Count <> 0 Then
            For slot = 1 To matchingSlots.Count
                If Npclist(NpcIndex).Invent.Object(matchingSlots.Item(slot)).Amount < MAX_INVENTORY_OBJS Then
                    SlotEnNPCInv = matchingSlots.Item(slot)
                    Exit Function
                End If
            Next slot
            SlotEnNPCInv = 0
            Exit Function
        End If
        
        SlotEnNPCInv = firstEmptySpace
        Exit Function

SlotEnNPCInv_Err:
120     Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.SlotEnNPCInv", Erl)
122     Resume Next
        
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
102     Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.Descuento", Erl)
104     Resume Next
        
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
        
100         Desc = Descuento(UserIndex)
        
            'Actualiza un solo slot
102         If Not UpdateAll Then
        
104             With Npclist(NpcIndex).Invent.Object(slot)
106                 obj.ObjIndex = .ObjIndex
108                 obj.Amount = .Amount
                
110                 If .ObjIndex > 0 Then
112                     val = (ObjData(.ObjIndex).Valor) / Desc
                    End If
                
114                 Call WriteChangeNPCInventorySlot(UserIndex, slot, obj, val)
                
                End With
            
            Else
        
                'Actualiza todos los slots
116             For LoopC = 1 To MAX_INVENTORY_SLOTS
            
118                 With Npclist(NpcIndex).Invent.Object(LoopC)
                
120                     obj.ObjIndex = .ObjIndex
122                     obj.Amount = .Amount
    
124                     If .ObjIndex > 0 Then
126                         val = (ObjData(.ObjIndex).Valor) / Desc
                        End If
    
128                     Call WriteChangeNPCInventorySlot(UserIndex, LoopC, obj, val)
                    
                    End With
                
130             Next LoopC
            
            End If

            Exit Sub

EnviarNpcInv_Err:
132         Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.UpdateNpcInv", Erl)
134         Resume Next
        
End Sub

''
' Update the inventory of the Npc to all users trading with him
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC
' @param slot The slot to update

Public Sub UpdateNpcInvToAll(ByVal UpdateAll As Boolean, ByVal NpcIndex As Integer, ByVal slot As Byte)
    '***************************************************
    On Error GoTo ErrHandler:

        Dim LoopC As Long

        ' Recorremos todos los usuarios
100     For LoopC = 1 To LastUser
    
102         With UserList(LoopC)
        
                ' Si esta comerciando
104             If .flags.Comerciando Then
            
                    ' Si el ultimo NPC que cliqueo es el que hay que actualizar
106                 If .flags.TargetNPC = NpcIndex Then
                        ' Actualizamos el inventario del NPC
108                     Call UpdateNpcInv(UpdateAll, LoopC, NpcIndex, slot)
                    End If
                
                End If
            
            End With
        
        Next
    
        Exit Sub
    
ErrHandler:
    
110     Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.UpdateNpcInvToAll")
112     Resume Next
    
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
106     Call RegistrarError(Err.Number, Err.Description, "modSistemaComercio.SalePrice", Erl)
108     Resume Next
        
End Function

