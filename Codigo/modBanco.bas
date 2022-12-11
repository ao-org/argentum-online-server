Attribute VB_Name = "modBanco"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
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

Sub IniciarDeposito(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler
        
        If UserList(UserIndex).flags.Comerciando Then
            Exit Sub
        End If
        
        UserList(UserIndex).flags.Comerciando = True
        
        Call UpdateBanUserInv(True, UserIndex, 0, "IniciarDeposito")

104     Call WriteBankInit(UserIndex)

    Exit Sub

ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modBanco.IniciarDeposito", Erl)

End Sub

Sub IniciarBanco(ByVal UserIndex As Integer)
        On Error GoTo IniciarBanco_Err

104     Call WriteGoliathInit(UserIndex)

        
        Exit Sub

IniciarBanco_Err:
106     Call TraceError(Err.Number, Err.Description, "modBanco.IniciarBanco", Erl)

        
End Sub

Sub SendBanObj(UserIndex As Integer, Slot As Byte, Object As t_UserOBJ)
        
        On Error GoTo SendBanObj_Err
        

100     UserList(UserIndex).BancoInvent.Object(Slot) = Object
        
102     Call WriteChangeBankSlot(UserIndex, Slot)

        
        Exit Sub

SendBanObj_Err:
104     Call TraceError(Err.Number, Err.Description, "modBanco.SendBanObj", Erl)

        
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte, caller As String)
        
        On Error GoTo UpdateBanUserInv_Err
        

        Dim NullObj As t_UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
            If Slot = 0 Then
                Exit Sub
            End If
102         If UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex > 0 Then
104             Call SendBanObj(UserIndex, Slot, UserList(UserIndex).BancoInvent.Object(Slot))
            Else
106             Call SendBanObj(UserIndex, Slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
110             If UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
112                 Call SendBanObj(UserIndex, LoopC, UserList(UserIndex).BancoInvent.Object(LoopC))
                Else
            
114                 Call SendBanObj(UserIndex, LoopC, NullObj)
            
                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateBanUserInv_Err:
118     Call TraceError(Err.Number, Err.Description + " UI: " & UserIndex & " Slot:" & Slot, "modBanco.UpdateBanUserInv " & "userID: " & UserList(UserIndex).ID & " Slot: " & Slot & " Caller: " & caller, Erl)

        
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)

        On Error GoTo ErrHandler

100     If Cantidad < 1 Then Exit Sub

        'Call WriteUpdateUserStats(UserIndex)
   
102     If UserList(UserIndex).BancoInvent.Object(i).amount > 0 Then
104         If Cantidad > UserList(UserIndex).BancoInvent.Object(i).amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(i).amount
            
            'Agregamos el obj que compro al inventario
106         slotdestino = UserReciveObj(UserIndex, CInt(i), Cantidad, slotdestino)
            
            If (slotdestino <> -1) Then
                'Actualizamos el inventario del usuario
108             Call UpdateUserInv(False, UserIndex, slotdestino)
    
                'Actualizamos el banco
110             Call UpdateBanUserInv(False, UserIndex, i, "UserRetiraItem")
            End If
            
        End If
    
        Exit Sub
    
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "modBanco.UsaRetiraItem")

    
End Sub

Function UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer) As Long
        
        On Error GoTo UserReciveObj_Err
    
        Dim Slot As Integer
        Dim obji As Integer

100     If UserList(UserIndex).BancoInvent.Object(ObjIndex).amount <= 0 Then Exit Function
    
        If (slotdestino > UserList(UserIndex).CurrentInventorySlots) Then ' Check exploit
            UserReciveObj = -1
            Exit Function
        End If

102     obji = UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex
        
        Dim slotvalido As Boolean

104     slotvalido = False

106     If slotdestino <> 0 Then

108         If UserList(UserIndex).Invent.Object(slotdestino).ObjIndex = 0 Then
110             slotvalido = True

            End If

            '¿Ya tiene un objeto de este tipo?

112         If UserList(UserIndex).Invent.Object(slotdestino).ObjIndex = obji And UserList(UserIndex).Invent.Object(slotdestino).amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
114             slotvalido = True

            End If

        End If

116     If slotvalido = False Then

            '¿Ya tiene un objeto de este tipo?
118         Slot = 1

120         Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji And UserList(UserIndex).Invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS
    
122             Slot = Slot + 1

124             If Slot > UserList(UserIndex).CurrentInventorySlots Then
                    Exit Do

                End If

            Loop

            'Sino se fija por un slot vacio
126         If Slot > UserList(UserIndex).CurrentInventorySlots Then
128             Slot = 1

130             Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
132                 Slot = Slot + 1

134                 If Slot > UserList(UserIndex).CurrentInventorySlots Then
136                     Call WriteConsoleMsg(UserIndex, "No podés tener mAs t_Objetos.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If

                Loop
138             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

            End If

        End If

140     If slotvalido Then
142         Slot = slotdestino
144         UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

        End If

        'Mete el obj en el slot
146     If UserList(UserIndex).Invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
            'Menor que MAX_INV_OBJS
148         UserList(UserIndex).Invent.Object(Slot).ObjIndex = obji
150         UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + Cantidad

            UserList(UserIndex).flags.ModificoInventario = True
    
152         Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Else
154         Call WriteConsoleMsg(UserIndex, "No podés tener mAs t_Objetos.", e_FontTypeNames.FONTTYPE_INFO)

        End If

        UserReciveObj = Slot
        
        Exit Function

UserReciveObj_Err:
156     Call TraceError(Err.Number, Err.Description, "modBanco.UserReciveObj", Erl)

        
End Function

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarBancoInvItem_Err
        
        Dim ObjIndex As Integer

100     ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex

        'Quita un Obj
102     UserList(UserIndex).BancoInvent.Object(Slot).amount = UserList(UserIndex).BancoInvent.Object(Slot).amount - Cantidad
        
104     If UserList(UserIndex).BancoInvent.Object(Slot).amount <= 0 Then
106         UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
108         UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
110         UserList(UserIndex).BancoInvent.Object(Slot).amount = 0

        End If
        
        UserList(UserIndex).flags.ModificoInventarioBanco = True
        
        Exit Sub

QuitarBancoInvItem_Err:
112     Call TraceError(Err.Number, Err.Description, "modBanco.QuitarBancoInvItem", Erl)

        
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)

        On Error GoTo ErrHandler

100     If UserList(UserIndex).Invent.Object(Item).amount > 0 And Cantidad > 0 Then
102         If Cantidad > UserList(UserIndex).Invent.Object(Item).amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).amount
        
            'Agregamos el obj que deposita al banco
104         slotdestino = UserDejaObj(UserIndex, CInt(Item), Cantidad, slotdestino)
        
            'Actualizamos el inventario del usuario
106         Call UpdateUserInv(False, UserIndex, Item)
        
            'Actualizamos el inventario del banco
108         Call UpdateBanUserInv(False, UserIndex, slotdestino, "UserDepositaItem")

        End If
    
        Exit Sub
    
ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "modBanco.UserDepositaItem")

    
End Sub

Sub UserDepositaItemDrop(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

        On Error GoTo ErrHandler

100     If UserList(UserIndex).Invent.Object(Item).amount > 0 And Cantidad > 0 Then
102         If Cantidad > UserList(UserIndex).Invent.Object(Item).amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).amount
            'Agregamos el obj que deposita al banco
104         Call UserDejaObj(UserIndex, CInt(Item), Cantidad, 0)
        
            'Actualizamos el inventario del usuario
106         Call UpdateUserInv(True, UserIndex, 0)
        
        End If
    
        Exit Sub

ErrHandler:
108     Call TraceError(Err.Number, Err.Description, "modBanco.UserDepositaItemDrop")


End Sub

Function UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer) As Long
        
        On Error GoTo UserDejaObj_Err
        
        Dim Slot As Integer
        Dim obji As Integer
    
100     If Cantidad < 1 Then Exit Function
102     obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
    
        Dim slotvalido As Boolean
104     slotvalido = False

106     If slotdestino <> 0 Then

108         If UserList(UserIndex).BancoInvent.Object(slotdestino).ObjIndex = 0 Then
110             slotvalido = True
            End If

112         If UserList(UserIndex).BancoInvent.Object(slotdestino).ObjIndex = obji And _
                UserList(UserIndex).BancoInvent.Object(slotdestino).amount + Cantidad <= MAX_INVENTORY_OBJS Then '¿Ya tiene un objeto de este tipo?
                
114             slotvalido = True
            End If

        End If

116     If slotvalido = False Then

            '¿Ya tiene un objeto de este tipo?
118         Slot = 1

120         Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji And _
                     UserList(UserIndex).BancoInvent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS
                     
122             Slot = Slot + 1
        
124             If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Exit Do
                End If
            Loop
    
            'Sino se fija por un slot vacio antes del slot devuelto
126         If Slot > MAX_BANCOINVENTORY_SLOTS Then
128             Slot = 1

130             Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
132                 Slot = Slot + 1
            
134                 If Slot > MAX_BANCOINVENTORY_SLOTS Then
136                     Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Function

                    End If

                Loop
        
138             UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
            End If

        End If
    
140     If slotvalido Then
142         Slot = slotdestino
144         UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1

            
        End If
        
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
    
            'Mete el obj en el slot
148         If UserList(UserIndex).BancoInvent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS Then
            
                'Menor que MAX_INV_OBJS
150             UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji
152             UserList(UserIndex).BancoInvent.Object(Slot).amount = UserList(UserIndex).BancoInvent.Object(Slot).amount + Cantidad
                
                UserList(UserIndex).flags.ModificoInventarioBanco = True
                
154             Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)

            Else
156             Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        UserDejaObj = Slot
        
        Exit Function

UserDejaObj_Err:
158     Call TraceError(Err.Number, Err.Description, "modBanco.UserDejaObj", Erl)

        
End Function

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserBovedaTxt_Err

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, e_FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, " Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", e_FontTypeNames.FONTTYPE_INFO)

104     For j = 1 To MAX_BANCOINVENTORY_SLOTS

106         If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).amount, e_FontTypeNames.FONTTYPE_INFO)

            End If

        Next

        
        Exit Sub

SendUserBovedaTxt_Err:
110     Call TraceError(Err.Number, Err.Description, "modBanco.SendUserBovedaTxt", Erl)

        
End Sub
