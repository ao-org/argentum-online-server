Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Sub IniciarDeposito(ByVal Userindex As Integer)

        On Error GoTo ErrHandler

        'Hacemos un Update del inventario del usuario
100     Call UpdateBanUserInv(True, Userindex, 0)
        'Actualizamos el dinero
102     Call WriteUpdateUserStats(Userindex)
        'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
104     Call WriteBankInit(Userindex)
106     UserList(Userindex).flags.Comerciando = True

ErrHandler:

End Sub

Sub IniciarBanco(ByVal Userindex As Integer)
        'Hacemos un Update del inventario del usuario
        
        On Error GoTo IniciarBanco_Err
        
100     Call UpdateBanUserInv(True, Userindex, 0)
        'Actualizamos el dinero
102     Call WriteUpdateUserStats(Userindex)

104     Call WriteGoliathInit(Userindex)

        
        Exit Sub

IniciarBanco_Err:
106     Call RegistrarError(Err.Number, Err.description, "modBanco.IniciarBanco", Erl)
108     Resume Next
        
End Sub

Sub SendBanObj(Userindex As Integer, slot As Byte, Object As UserOBJ)
        
        On Error GoTo SendBanObj_Err
        

100     UserList(Userindex).BancoInvent.Object(slot) = Object

102     Call WriteChangeBankSlot(Userindex, slot)

        
        Exit Sub

SendBanObj_Err:
104     Call RegistrarError(Err.Number, Err.description, "modBanco.SendBanObj", Erl)
106     Resume Next
        
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateBanUserInv_Err
        

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(Userindex).BancoInvent.Object(slot).ObjIndex > 0 Then
104             Call SendBanObj(Userindex, slot, UserList(Userindex).BancoInvent.Object(slot))
            Else
106             Call SendBanObj(Userindex, slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
110             If UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
112                 Call SendBanObj(Userindex, LoopC, UserList(Userindex).BancoInvent.Object(LoopC))
                Else
            
114                 Call SendBanObj(Userindex, LoopC, NullObj)
            
                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateBanUserInv_Err:
118     Call RegistrarError(Err.Number, Err.description, "modBanco.UpdateBanUserInv", Erl)
120     Resume Next
        
End Sub

Sub UserRetiraItem(ByVal Userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)

        On Error GoTo ErrHandler

100     If Cantidad < 1 Then Exit Sub

        'Call WriteUpdateUserStats(UserIndex)
   
102     If UserList(Userindex).BancoInvent.Object(i).Amount > 0 Then
104         If Cantidad > UserList(Userindex).BancoInvent.Object(i).Amount Then Cantidad = UserList(Userindex).BancoInvent.Object(i).Amount
            
            'Agregamos el obj que compro al inventario
106         Call UserReciveObj(Userindex, CInt(i), Cantidad, slotdestino)
            'Actualizamos el inventario del usuario
108         Call UpdateUserInv(True, Userindex, 0)
            'Actualizamos el banco
110         Call UpdateBanUserInv(True, Userindex, 0)

        End If
    
        Exit Sub
    
ErrHandler:
112     Call RegistrarError(Err.Number, Err.description, "modBanco.UsaRetiraItem")
114     Resume Next
    
End Sub

Sub UserReciveObj(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)
        
        On Error GoTo UserReciveObj_Err
        

        Dim slot As Integer

        Dim obji As Integer

100     If UserList(Userindex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

102     obji = UserList(Userindex).BancoInvent.Object(ObjIndex).ObjIndex

        Dim slotvalido As Boolean

104     slotvalido = False

106     If slotdestino <> 0 Then

108         If UserList(Userindex).Invent.Object(slotdestino).ObjIndex = 0 Then
110             slotvalido = True

            End If

            '¿Ya tiene un objeto de este tipo?

112         If UserList(Userindex).Invent.Object(slotdestino).ObjIndex = obji And UserList(Userindex).Invent.Object(slotdestino).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
114             slotvalido = True

            End If

        End If

116     If slotvalido = False Then

            '¿Ya tiene un objeto de este tipo?
118         slot = 1

120         Do Until UserList(Userindex).Invent.Object(slot).ObjIndex = obji And UserList(Userindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
122             slot = slot + 1

124             If slot > UserList(Userindex).CurrentInventorySlots Then
                    Exit Do

                End If

            Loop

            'Sino se fija por un slot vacio
126         If slot > UserList(Userindex).CurrentInventorySlots Then
128             slot = 1

130             Do Until UserList(Userindex).Invent.Object(slot).ObjIndex = 0
132                 slot = slot + 1

134                 If slot > UserList(Userindex).CurrentInventorySlots Then
136                     Call WriteConsoleMsg(Userindex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                Loop
138             UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems + 1

            End If

        End If

140     If slotvalido Then
142         slot = slotdestino
144         UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems + 1

        End If

        'Mete el obj en el slot
146     If UserList(Userindex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
            'Menor que MAX_INV_OBJS
148         UserList(Userindex).Invent.Object(slot).ObjIndex = obji
150         UserList(Userindex).Invent.Object(slot).Amount = UserList(Userindex).Invent.Object(slot).Amount + Cantidad
    
152         Call QuitarBancoInvItem(Userindex, CByte(ObjIndex), Cantidad)
        Else
154         Call WriteConsoleMsg(Userindex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

UserReciveObj_Err:
156     Call RegistrarError(Err.Number, Err.description, "modBanco.UserReciveObj", Erl)
158     Resume Next
        
End Sub

Sub QuitarBancoInvItem(ByVal Userindex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarBancoInvItem_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = UserList(Userindex).BancoInvent.Object(slot).ObjIndex

        'Quita un Obj

102     UserList(Userindex).BancoInvent.Object(slot).Amount = UserList(Userindex).BancoInvent.Object(slot).Amount - Cantidad
        
104     If UserList(Userindex).BancoInvent.Object(slot).Amount <= 0 Then
106         UserList(Userindex).BancoInvent.NroItems = UserList(Userindex).BancoInvent.NroItems - 1
108         UserList(Userindex).BancoInvent.Object(slot).ObjIndex = 0
110         UserList(Userindex).BancoInvent.Object(slot).Amount = 0

        End If
    
        
        Exit Sub

QuitarBancoInvItem_Err:
112     Call RegistrarError(Err.Number, Err.description, "modBanco.QuitarBancoInvItem", Erl)
114     Resume Next
        
End Sub

Sub UserDepositaItem(ByVal Userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)

        On Error GoTo ErrHandler

100     If UserList(Userindex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
102         If Cantidad > UserList(Userindex).Invent.Object(Item).Amount Then Cantidad = UserList(Userindex).Invent.Object(Item).Amount
        
            'Agregamos el obj que deposita al banco
104         Call UserDejaObj(Userindex, CInt(Item), Cantidad, slotdestino)
        
            'Actualizamos el inventario del usuario
106         Call UpdateUserInv(True, Userindex, 0)
        
            'Actualizamos el inventario del banco
108         Call UpdateBanUserInv(True, Userindex, 0)

        End If
    
        Exit Sub
    
ErrHandler:
110     Call RegistrarError(Err.Number, Err.description, "modBanco.UserDepositaItem")
112     Resume Next
    
End Sub

Sub UserDepositaItemDrop(ByVal Userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

        On Error GoTo ErrHandler

100     If UserList(Userindex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
102         If Cantidad > UserList(Userindex).Invent.Object(Item).Amount Then Cantidad = UserList(Userindex).Invent.Object(Item).Amount
            'Agregamos el obj que deposita al banco
104         Call UserDejaObj(Userindex, CInt(Item), Cantidad, 0)
        
            'Actualizamos el inventario del usuario
106         Call UpdateUserInv(True, Userindex, 0)
        
        End If
    
        Exit Sub

ErrHandler:
108     Call RegistrarError(Err.Number, Err.description, "modBanco.UserDepositaItemDrop")
110     Resume Next

End Sub

Sub UserDejaObj(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)
        
        On Error GoTo UserDejaObj_Err
        

        Dim slot As Integer

        Dim obji As Integer
    
100     If Cantidad < 1 Then Exit Sub
102     obji = UserList(Userindex).Invent.Object(ObjIndex).ObjIndex
    
        Dim slotvalido As Boolean

104     slotvalido = False

106     If slotdestino <> 0 Then
108         If UserList(Userindex).BancoInvent.Object(slotdestino).ObjIndex = 0 Then
110             slotvalido = True

            End If

112         If UserList(Userindex).BancoInvent.Object(slotdestino).ObjIndex = obji And UserList(Userindex).BancoInvent.Object(slotdestino).Amount + Cantidad <= MAX_INVENTORY_OBJS Then '¿Ya tiene un objeto de este tipo?
114             slotvalido = True

            End If

        End If

116     If slotvalido = False Then
            '¿Ya tiene un objeto de este tipo?
118         slot = 1

120         Do Until UserList(Userindex).BancoInvent.Object(slot).ObjIndex = obji And UserList(Userindex).BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
122             slot = slot + 1
        
124             If slot > MAX_BANCOINVENTORY_SLOTS Then
                    Exit Do

                End If

            Loop
    
            'Sino se fija por un slot vacio antes del slot devuelto
126         If slot > MAX_BANCOINVENTORY_SLOTS Then
128             slot = 1

130             Do Until UserList(Userindex).BancoInvent.Object(slot).ObjIndex = 0
132                 slot = slot + 1
            
134                 If slot > MAX_BANCOINVENTORY_SLOTS Then
136                     Call WriteConsoleMsg(Userindex, "No tienes mas espacio en el banco.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

                Loop
        
138             UserList(Userindex).BancoInvent.NroItems = UserList(Userindex).BancoInvent.NroItems + 1

            End If

        End If
    
140     If slotvalido Then
142         slot = slotdestino
144         UserList(Userindex).BancoInvent.NroItems = UserList(Userindex).BancoInvent.NroItems + 1

        End If
    
146     If slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
148         If UserList(Userindex).BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            
                'Menor que MAX_INV_OBJS
150             UserList(Userindex).BancoInvent.Object(slot).ObjIndex = obji
152             UserList(Userindex).BancoInvent.Object(slot).Amount = UserList(Userindex).BancoInvent.Object(slot).Amount + Cantidad
            
154             Call QuitarUserInvItem(Userindex, CByte(ObjIndex), Cantidad)
            Else
156             Call WriteConsoleMsg(Userindex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

UserDejaObj_Err:
158     Call RegistrarError(Err.Number, Err.description, "modBanco.UserDejaObj", Erl)
160     Resume Next
        
End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)

        On Error Resume Next

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(Userindex).name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, " Tiene " & UserList(Userindex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

104     For j = 1 To MAX_BANCOINVENTORY_SLOTS

106         If UserList(Userindex).BancoInvent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(Userindex).BancoInvent.Object(j).ObjIndex).name & " Cantidad:" & UserList(Userindex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

            End If

        Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

        On Error Resume Next

        Dim j        As Integer

        Dim CharFile As String, Tmp As String

        Dim ObjInd   As Long, ObjCant As Long

100     CharFile = CharPath & CharName & ".chr"

102     If FileExist(CharFile, vbNormal) Then
104         Call WriteConsoleMsg(sendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

108         For j = 1 To MAX_BANCOINVENTORY_SLOTS
110             Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
112             ObjInd = ReadField(1, Tmp, Asc("-"))
114             ObjCant = ReadField(2, Tmp, Asc("-"))

116             If ObjInd > 0 Then
118                 Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

                End If

            Next
        Else
120         Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)

        End If

End Sub

