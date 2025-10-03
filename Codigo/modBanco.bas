Attribute VB_Name = "modBanco"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Sub IniciarDeposito(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If UserList(UserIndex).flags.Comerciando Then
        Exit Sub
    End If
    UserList(UserIndex).flags.Comerciando = True
    Call UpdateBanUserInv(True, UserIndex, 0, "IniciarDeposito")
    Call WriteBankInit(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modBanco.IniciarDeposito", Erl)
End Sub

Sub IniciarBanco(ByVal UserIndex As Integer)
    On Error GoTo IniciarBanco_Err
    Call WriteGoliathInit(UserIndex)
    Exit Sub
IniciarBanco_Err:
    Call TraceError(Err.Number, Err.Description, "modBanco.IniciarBanco", Erl)
End Sub

Sub SendBanObj(UserIndex As Integer, Slot As Byte, Object As t_UserOBJ)
    On Error GoTo SendBanObj_Err
    UserList(UserIndex).BancoInvent.Object(Slot) = Object
    Call WriteChangeBankSlot(UserIndex, Slot)
    Exit Sub
SendBanObj_Err:
    Call TraceError(Err.Number, Err.Description, "modBanco.SendBanObj", Erl)
End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte, caller As String)
    On Error GoTo UpdateBanUserInv_Err
    Dim NullObj As t_UserOBJ
    Dim LoopC   As Byte
    'Actualiza un solo slot
    If Not UpdateAll Then
        'Actualiza el inventario
        If Slot = 0 Then
            Exit Sub
        End If
        If UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex > 0 Then
            Call SendBanObj(UserIndex, Slot, UserList(UserIndex).BancoInvent.Object(Slot))
        Else
            Call SendBanObj(UserIndex, Slot, NullObj)
        End If
    Else
        'Actualiza todos los slots
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            'Actualiza el inventario
            If UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex > 0 Then
                Call SendBanObj(UserIndex, LoopC, UserList(UserIndex).BancoInvent.Object(LoopC))
            Else
                Call SendBanObj(UserIndex, LoopC, NullObj)
            End If
        Next LoopC
    End If
    Exit Sub
UpdateBanUserInv_Err:
    Call TraceError(Err.Number, Err.Description + " UI: " & UserIndex & " Slot:" & Slot, "modBanco.UpdateBanUserInv " & "userID: " & UserList(UserIndex).Id & " Slot: " & Slot & _
            " Caller: " & caller, Erl)
End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)
    On Error GoTo ErrHandler
    If Cantidad < 1 Then Exit Sub
    'Call WriteUpdateUserStats(UserIndex)
    If UserList(UserIndex).BancoInvent.Object(i).amount > 0 Then
        If Cantidad > UserList(UserIndex).BancoInvent.Object(i).amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(i).amount
        'Agregamos el obj que compro al inventario
        slotdestino = UserReciveObj(UserIndex, CInt(i), Cantidad, slotdestino)
        If (slotdestino <> -1) Then
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, UserIndex, slotdestino)
            'Actualizamos el banco
            Call UpdateBanUserInv(False, UserIndex, i, "UserRetiraItem")
        End If
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modBanco.UsaRetiraItem")
End Sub

Function UserReciveObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer) As Long
    On Error GoTo UserReciveObj_Err
    Dim Slot As Integer
    Dim obji As Integer
    If UserList(UserIndex).BancoInvent.Object(ObjIndex).amount <= 0 Then Exit Function
    If (slotdestino > UserList(UserIndex).CurrentInventorySlots) Then ' Check exploit
        UserReciveObj = -1
        Exit Function
    End If
    obji = UserList(UserIndex).BancoInvent.Object(ObjIndex).ObjIndex
    Dim slotvalido As Boolean
    slotvalido = False
    If slotdestino <> 0 Then
        If UserList(UserIndex).invent.Object(slotdestino).ObjIndex = 0 Then
            slotvalido = True
        End If
        '¿Ya tiene un objeto de este tipo?
        If UserList(UserIndex).invent.Object(slotdestino).ObjIndex = obji And UserList(UserIndex).invent.Object(slotdestino).amount + Cantidad <= MAX_INVENTORY_OBJS And UserList( _
                UserIndex).invent.Object(slotdestino).ElementalTags = UserList(UserIndex).BancoInvent.Object(ObjIndex).ElementalTags Then
            slotvalido = True
        End If
    End If
    If slotvalido = False Then
        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until UserList(UserIndex).invent.Object(Slot).ObjIndex = obji And UserList(UserIndex).invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS And UserList( _
                UserIndex).invent.Object(Slot).ElementalTags = UserList(UserIndex).BancoInvent.Object(ObjIndex).ElementalTags
            Slot = Slot + 1
            If Slot > UserList(UserIndex).CurrentInventorySlots Then
                Exit Do
            End If
        Loop
        'Sino se fija por un slot vacio
        If Slot > UserList(UserIndex).CurrentInventorySlots Then
            Slot = 1
            Do Until UserList(UserIndex).invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                If Slot > UserList(UserIndex).CurrentInventorySlots Then
                    Call WriteLocaleMsg(UserIndex, 1600, e_FontTypeNames.FONTTYPE_INFO) 'Msg1600= No podés tener más objetos.
                    Exit Function
                End If
            Loop
            UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems + 1
        End If
    End If
    If slotvalido Then
        Slot = slotdestino
        UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems + 1
    End If
    'Mete el obj en el slot
    If UserList(UserIndex).invent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).invent.Object(Slot).ObjIndex = obji
        UserList(UserIndex).invent.Object(Slot).amount = UserList(UserIndex).invent.Object(Slot).amount + Cantidad
        UserList(UserIndex).invent.Object(Slot).ElementalTags = UserList(UserIndex).BancoInvent.Object(ObjIndex).ElementalTags
        UserList(UserIndex).flags.ModificoInventario = True
        Call QuitarBancoInvItem(UserIndex, CByte(ObjIndex), Cantidad)
    Else
        Call WriteLocaleMsg(UserIndex, 1600, e_FontTypeNames.FONTTYPE_INFO) 'Msg1600= No podés tener más objetos.
    End If
    UserReciveObj = Slot
    Exit Function
UserReciveObj_Err:
    Call TraceError(Err.Number, Err.Description, "modBanco.UserReciveObj", Erl)
End Function

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    On Error GoTo QuitarBancoInvItem_Err
    With UserList(UserIndex).BancoInvent
        'Quita un Obj
        .Object(Slot).amount = .Object(Slot).amount - Cantidad
        If .Object(Slot).amount <= 0 Then
            .NroItems = .NroItems - 1
            .Object(Slot).ObjIndex = 0
            .Object(Slot).amount = 0
            .Object(Slot).ElementalTags = 0
        End If
        UserList(UserIndex).flags.ModificoInventarioBanco = True
    End With
    Exit Sub
QuitarBancoInvItem_Err:
    Call TraceError(Err.Number, Err.Description, "modBanco.QuitarBancoInvItem", Erl)
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer)
    On Error GoTo ErrHandler
    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then
        Exit Sub
    End If
    If UserList(UserIndex).invent.Object(Slot).amount > 0 And Cantidad > 0 Then
        If Cantidad > UserList(UserIndex).invent.Object(Slot).amount Then Cantidad = UserList(UserIndex).invent.Object(Slot).amount
        'Agregamos el obj que deposita al banco
        slotdestino = UserDejaObj(UserIndex, CInt(Slot), Cantidad, slotdestino)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(False, UserIndex, Slot)
        'Actualizamos el inventario del banco
        Call UpdateBanUserInv(False, UserIndex, slotdestino, "UserDepositaItem")
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modBanco.UserDepositaItem")
End Sub

Function UserDejaObj(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, ByVal slotdestino As Integer) As Long
    On Error GoTo UserDejaObj_Err
    Dim Slot As Integer
    Dim obji As Integer
    If Cantidad < 1 Then Exit Function
    obji = UserList(UserIndex).invent.Object(ObjIndex).ObjIndex
    Dim slotvalido As Boolean
    slotvalido = False
    If slotdestino <> 0 Then
        If UserList(UserIndex).BancoInvent.Object(slotdestino).ObjIndex = 0 Then
            slotvalido = True
        End If
        If UserList(UserIndex).BancoInvent.Object(slotdestino).ObjIndex = obji And UserList(UserIndex).BancoInvent.Object(slotdestino).ElementalTags = UserList( _
                UserIndex).invent.Object(ObjIndex).ElementalTags And UserList(UserIndex).BancoInvent.Object(slotdestino).amount + Cantidad <= MAX_INVENTORY_OBJS Then '¿Ya tiene un objeto de este tipo?
            slotvalido = True
        End If
    End If
    If slotvalido = False Then
        '¿Ya tiene un objeto de este tipo?
        Slot = 1
        Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji And UserList(UserIndex).BancoInvent.Object(Slot).ElementalTags = UserList(UserIndex).invent.Object( _
                ObjIndex).ElementalTags And UserList(UserIndex).BancoInvent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
        Loop
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1
            Do Until UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteLocaleMsg(UserIndex, 1601, e_FontTypeNames.FONTTYPE_INFOIAO) 'Msg1601= No tienes más espacio en el banco.
                    Exit Function
                End If
            Loop
            UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
        End If
    End If
    If slotvalido Then
        Slot = slotdestino
        UserList(UserIndex).BancoInvent.NroItems = UserList(UserIndex).BancoInvent.NroItems + 1
    End If
    If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        If UserList(UserIndex).BancoInvent.Object(Slot).amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex = obji
            UserList(UserIndex).BancoInvent.Object(Slot).amount = UserList(UserIndex).BancoInvent.Object(Slot).amount + Cantidad
            UserList(UserIndex).BancoInvent.Object(Slot).ElementalTags = UserList(UserIndex).invent.Object(ObjIndex).ElementalTags
            UserList(UserIndex).flags.ModificoInventarioBanco = True
            Call QuitarUserInvItem(UserIndex, CByte(ObjIndex), Cantidad)
        Else
            Call WriteLocaleMsg(UserIndex, 1602, e_FontTypeNames.FONTTYPE_INFO) 'Msg1602= El banco no puede cargar tantos objetos.
        End If
    End If
    UserDejaObj = Slot
    Exit Function
UserDejaObj_Err:
    Call TraceError(Err.Number, Err.Description, "modBanco.UserDejaObj", Erl)
End Function

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo SendUserBovedaTxt_Err
    Dim j As Integer
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1939, UserList(UserIndex).BancoInvent.NroItems, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1939= Tiene ¬1 objetos.
    For j = 1 To MAX_BANCOINVENTORY_SLOTS
        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1940, j & "¬" & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).name & "¬" & UserList( _
                    UserIndex).BancoInvent.Object(j).amount, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1940= Objeto ¬1 ¬2 Cantidad:¬3
        End If
    Next
    Exit Sub
SendUserBovedaTxt_Err:
    Call TraceError(Err.Number, Err.Description, "modBanco.SendUserBovedaTxt", Erl)
End Sub
