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

Sub IniciarDeposito(ByVal UserIndex As Integer)

    On Error GoTo Errhandler

    'Hacemos un Update del inventario del usuario
    Call UpdateBanUserInv(True, UserIndex, 0)
    'Actualizamos el dinero
    Call WriteUpdateUserStats(UserIndex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    Call WriteBankInit(UserIndex)
    UserList(UserIndex).flags.Comerciando = True

Errhandler:

End Sub

Sub IniciarBanco(ByVal UserIndex As Integer)
        'Hacemos un Update del inventario del usuario
        
        On Error GoTo IniciarBanco_Err
        
100     Call UpdateBanUserInv(True, UserIndex, 0)
        'Actualizamos el dinero
102     Call WriteUpdateUserStats(UserIndex)

104     Call WriteGoliathInit(UserIndex)

        
        Exit Sub

IniciarBanco_Err:
        Call RegistrarError(Err.Number, Err.description, "modBanco.IniciarBanco", Erl)
        Resume Next
        
End Sub

Sub SendBanObj(UserIndex As Integer, slot As Byte, Object As UserOBJ)
        
        On Error GoTo SendBanObj_Err
        

100     UserList(UserIndex).BancoInvent.Object(slot) = Object

102     Call WriteChangeBankSlot(UserIndex, slot)

        
        Exit Sub

SendBanObj_Err:
        Call RegistrarError(Err.Number, Err.description, "modBanco.SendBanObj", Erl)
        Resume Next
        
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateBanUserInv_Err
        

        Dim NullObj As UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(UserIndex).BancoInvent.Object(slot).ObjIndex > 0 Then
104             Call SendBanObj(UserIndex, slot, UserList(UserIndex).BancoInvent.Object(slot))
            Else
106             Call SendBanObj(UserIndex, slot, NullObj)

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
        Call RegistrarError(Err.Number, Err.description, "modBanco.UpdateBanUserInv", Erl)
        Resume Next
        
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)

    On Error GoTo Errhandler

    If Cantidad < 1 Then Exit Sub

    'Call WriteUpdateUserStats(UserIndex)
   
    If UserList(UserIndex).BancoInvent.Object(i).Amount > 0 Then
        If Cantidad > UserList(UserIndex).BancoInvent.Object(i).Amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(i).Amount
            
        'Agregamos el obj que compro al inventario
        Call UserReciveObj(UserIndex, CInt(i), Cantidad, slotdestino)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        'Actualizamos el banco
        Call UpdateBanUserInv(True, UserIndex, 0)

    End If
    
    Exit Sub
    
Errhandler:
    Call RegistrarError(Err.Number, Err.description, "modBanco.UsaRetiraItem")
    Resume Next
    
End Sub

Sub UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)
        
        On Error GoTo UserReciveObj_Err
        

        Dim slot As Integer

        Dim obji As Integer

100     If UserList(UserIndex).BancoInvent.Object(ObjIndex).Amount <= 0 Then Exit Sub

102     obji = UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex

        Dim slotvalido As Boolean

104     slotvalido = False

106     If slotdestino <> 0 Then

108         If UserList(UserIndex).Invent.Object(slotdestino).ObjIndex = 0 Then
110             slotvalido = True

            End If

            '�Ya tiene un objeto de este tipo?

112         If UserList(UserIndex).Invent.Object(slotdestino).ObjIndex = obji And UserList(UserIndex).Invent.Object(slotdestino).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
114             slotvalido = True

            End If

        End If

116     If slotvalido = False Then

            '�Ya tiene un objeto de este tipo?
118         slot = 1

120         Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = obji And UserList(UserIndex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
    
122             slot = slot + 1

124             If slot > UserList(UserIndex).CurrentInventorySlots Then
                    Exit Do

                End If

            Loop

            'Sino se fija por un slot vacio
126         If slot > UserList(UserIndex).CurrentInventorySlots Then
128             slot = 1

130             Do Until UserList(UserIndex).Invent.Object(slot).ObjIndex = 0
132                 slot = slot + 1

134                 If slot > UserList(UserIndex).CurrentInventorySlots Then
136                     Call WriteConsoleMsg(UserIndex, "No pod�s tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                Loop
138             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

            End If

        End If

140     If slotvalido Then
142         slot = slotdestino
144         UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1

        End If

        'Mete el obj en el slot
146     If UserList(UserIndex).Invent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
    
            'Menor que MAX_INV_OBJS
148         UserList(UserIndex).Invent.Object(slot).ObjIndex = obji
150         UserList(UserIndex).Invent.Object(slot).Amount = UserList(UserIndex).Invent.Object(slot).Amount + Cantidad
    
152         Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Else
154         Call WriteConsoleMsg(UserIndex, "No pod�s tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

UserReciveObj_Err:
        Call RegistrarError(Err.Number, Err.description, "modBanco.UserReciveObj", Erl)
        Resume Next
        
End Sub

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarBancoInvItem_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = UserList(UserIndex).BancoInvent.Object(slot).ObjIndex

        'Quita un Obj

102     UserList(UserIndex).BancoInvent.Object(slot).Amount = UserList(UserIndex).BancoInvent.Object(slot).Amount - Cantidad
        
104     If UserList(UserIndex).BancoInvent.Object(slot).Amount <= 0 Then
106         UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems - 1
108         UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = 0
110         UserList(UserIndex).BancoInvent.Object(slot).Amount = 0

        End If
    
        
        Exit Sub

QuitarBancoInvItem_Err:
        Call RegistrarError(Err.Number, Err.description, "modBanco.QuitarBancoInvItem", Erl)
        Resume Next
        
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)

    On Error GoTo Errhandler

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
        If Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        
        'Agregamos el obj que deposita al banco
        Call UserDejaObj(UserIndex, CInt(Item), Cantidad, slotdestino)
        
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        
        'Actualizamos el inventario del banco
        Call UpdateBanUserInv(True, UserIndex, 0)

    End If
    
Errhandler:
    Call RegistrarError(Err.Number, Err.description, "modBanco.UserDepositaItem")
    Resume Next
    
End Sub

Sub UserDepositaItemDrop(ByVal UserIndex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)

    On Error GoTo Errhandler

    If UserList(UserIndex).Invent.Object(Item).Amount > 0 And Cantidad > 0 Then
        If Cantidad > UserList(UserIndex).Invent.Object(Item).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Item).Amount
        'Agregamos el obj que deposita al banco
        Call UserDejaObj(UserIndex, CInt(Item), Cantidad, 0)
        
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, UserIndex, 0)
        
    End If

Errhandler:
    Call RegistrarError(Err.Number, Err.description, "modBanco.UserDepositaItemDrop")
    Resume Next

End Sub

Sub UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)
        
        On Error GoTo UserDejaObj_Err
        

        Dim slot As Integer

        Dim obji As Integer
    
100     If Cantidad < 1 Then Exit Sub
102     obji = UserList(UserIndex).Invent.Object(ObjIndex).ObjIndex
    
        Dim slotvalido As Boolean

104     slotvalido = False

106     If slotdestino <> 0 Then
108         If UserList(UserIndex).BancoInvent.Object(slotdestino).ObjIndex = 0 Then
110             slotvalido = True

            End If

112         If UserList(UserIndex).BancoInvent.Object(slotdestino).ObjIndex = obji And UserList(UserIndex).BancoInvent.Object(slotdestino).Amount + Cantidad <= MAX_INVENTORY_OBJS Then '�Ya tiene un objeto de este tipo?
114             slotvalido = True

            End If

        End If

116     If slotvalido = False Then
            '�Ya tiene un objeto de este tipo?
118         slot = 1

120         Do Until UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = obji And UserList(UserIndex).BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
122             slot = slot + 1
        
124             If slot > MAX_BANCOINVENTORY_SLOTS Then
                    Exit Do

                End If

            Loop
    
            'Sino se fija por un slot vacio antes del slot devuelto
126         If slot > MAX_BANCOINVENTORY_SLOTS Then
128             slot = 1

130             Do Until UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = 0
132                 slot = slot + 1
            
134                 If slot > MAX_BANCOINVENTORY_SLOTS Then
136                     Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

                Loop
        
138             UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1

            End If

        End If
    
140     If slotvalido Then
142         slot = slotdestino
144         UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1

        End If
    
146     If slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
148         If UserList(UserIndex).BancoInvent.Object(slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            
                'Menor que MAX_INV_OBJS
150             UserList(UserIndex).BancoInvent.Object(slot).ObjIndex = obji
152             UserList(UserIndex).BancoInvent.Object(slot).Amount = UserList(UserIndex).BancoInvent.Object(slot).Amount + Cantidad
            
154             Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
            Else
156             Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

UserDejaObj_Err:
        Call RegistrarError(Err.Number, Err.description, "modBanco.UserDejaObj", Erl)
        Resume Next
        
End Sub

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)

    On Error Resume Next

    Dim j As Integer

    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, " Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

        End If

    Next

End Sub

Sub SendUserBovedaTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)

    On Error Resume Next

    Dim j        As Integer

    Dim CharFile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long

    CharFile = CharPath & CharName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, CharName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & CharName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

