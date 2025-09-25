Attribute VB_Name = "ModSubasta"

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
Public Type t_Subastas
    HaySubastaActiva As Boolean
    SubastaHabilitada As Boolean
    ObjSubastado As Integer
    ObjSubastadoCantidad As Integer
    OfertaInicial As Long
    Subastador As String
    MejorOferta As Long
    Comprador As String
    HuboOferta As Boolean
    TiempoRestanteSubasta As Integer
    MinutosDeSubasta As Byte
    PosibleCancelo As Boolean
End Type

Public Subasta As t_Subastas
Dim Logear     As String

Public Sub IniciarSubasta(ByVal UserIndex As Integer)
    On Error GoTo IniciarSubasta_Err
    If UserList(UserIndex).flags.Subastando = True And Not Subasta.HaySubastaActiva Then
        Call WriteLocaleChatOverHead(UserIndex, 1427, UserList(UserIndex).Counters.TiempoParaSubastar, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, _
                vbWhite) ' Msg1427=Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. Te quedan: ¬1 segundos... ¡Apurate!
        Exit Sub
    End If
    If Subasta.HaySubastaActiva = True Then
        Call WriteLocaleChatOverHead(UserIndex, 1428, "", str$(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1428=Oye amigo, espera tu turno, estoy subastando en este momento.
        Exit Sub
    End If
    If Not MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex > 0 Then
        Call WriteLocaleChatOverHead(UserIndex, 1429, "", str$(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1429=¿Pues Acaso el aire está en venta ahora? ¡Bribón!
        Exit Sub
    End If
    If Not ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex).Subastable = 1 Then
        Call WriteLocaleChatOverHead(UserIndex, 1430, "", str$(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1430=Aquí solo subastamos items que sean valiosos. ¡Largate de acá Bribón!
        Exit Sub
    End If
    If UserList(UserIndex).flags.Subastando = True Then 'Practicamente imposible que pase... pero por si las dudas
        Call WriteLocaleChatOverHead(UserIndex, 1431, "", str$(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex), vbRed) ' Msg1431=Tu ya estas subastando! Esto a quedado logeado.
        Logear = "El usuario que ya estaba subastando pudo subastar otro item" & Date & " - " & Time
        Call LogearEventoDeSubasta(Logear)
        Exit Sub
    End If
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex > 0 Then
        Subasta.ObjSubastado = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex
        Subasta.ObjSubastadoCantidad = MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.amount
        Subasta.Subastador = UserList(UserIndex).name
        UserList(UserIndex).Counters.TiempoParaSubastar = 15
        Call WriteLocaleChatOverHead(UserIndex, 1432, "", str$(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1432=Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. ¡Tienes 15 segundos!
        Call EraseObj(Subasta.ObjSubastadoCantidad, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
        UserList(UserIndex).flags.Subastando = True
        Exit Sub
    End If
    Exit Sub
IniciarSubasta_Err:
    Call TraceError(Err.Number, Err.Description, "ModSubasta.IniciarSubasta", Erl)
End Sub

Public Sub FinalizarSubasta()
    'Primero Damos el objeto subastado
    'Despues el oro al subastador
    On Error GoTo FinalizarSubasta_Err
    '1) nos fijamos si el usuario que gano la subasta esta online,
    'si esta online le ponemos el objeto en el inventario, si no tiene
    'lugar en el inventario lo tiramos al piso,
    'si esta offline el item se va a depositar en boveda.
    'Si no tiene lugar en boveda, el item se tira en la ultima posicion
    'donde el user estubo parado.
    '2)Nos fijamos si esta online, si esta online, le damos el oro a la billetera,
    'si esta offline se deposita en el banco.
    'El sistema le cobra un 10% del precio de venta, por uso de servicio.
    Dim ObjVendido   As t_Obj
    Dim tUser        As t_UserReference
    Dim Subastador   As t_UserReference
    Dim Leer         As New clsIniManager
    Dim FileUser     As String
    Dim SlotEnBoveda As Integer
    Dim PosMap       As Byte
    Dim PosX         As Byte
    Dim PosY         As Byte
    FileUser = CharPath & UCase$(Subasta.Comprador) & ".chr"
    ObjVendido.ObjIndex = Subasta.ObjSubastado
    ObjVendido.amount = Subasta.ObjSubastadoCantidad
    tUser = NameIndex(Subasta.Comprador)
    If Not IsValidUserRef(tUser) Then
        Call LogearEventoDeSubasta("El usuario ganador de subasta se encuentra offline, intentando depositar en boveda")
        Call Leer.Initialize(FileUser)
        SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1
        If SlotEnBoveda < MAX_BANCOINVENTORY_SLOTS Then
            Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
            Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te deposite el item en la boveda.")
            Call LogearEventoDeSubasta("El items fue depositado en la boveda del comprador correctamente.")
        Else
            PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
            PosX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
            PosY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))
            If MapData(PosMap, PosX, PosY).ObjInfo.ObjIndex > 0 Then Exit Sub
            If MapData(PosMap, PosX, PosY).TileExit.Map > 0 Then Exit Sub
            If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub
            If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
            Call MakeObj(ObjVendido, PosMap, PosX, PosY)
            Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se tiro en la posicion:" & PosMap & "-" & PosX & "-" & PosY)
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", _
                    "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Como no tenias espacio ni en tu boveda ni en el correo, tuve que tirarlo en tu ultima posicion.")
        End If
    Else
        If Not MeterItemEnInventario(tUser.ArrayIndex, ObjVendido) Then
            Call TirarItemAlPiso(UserList(tUser.ArrayIndex).pos, ObjVendido)
        End If
        Call LogearEventoDeSubasta("Se entrego el item en mano.")
        Call WriteLocaleMsg(tUser.ArrayIndex, "1440", e_FontTypeNames.FONTTYPE_SUBASTA) ' Msg1440=Felicitaciones, has ganado la subasta.
    End If
    Dim Descuento As Long
    Descuento = Subasta.MejorOferta / 100 * 5
    Subasta.MejorOferta = Subasta.MejorOferta - Descuento
    Subastador = NameIndex(Subasta.Subastador)
    If Not IsValidUserRef(Subastador) Then
        Call LogearEventoDeSubasta("El subastador se encontraba offline cuando se le tenia que dar el oro, depositando en el banco.")
        Call Leer.Initialize(CharPath & UCase$(Subasta.Subastador) & ".chr")
        Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "STATS", "Banco", CLng(Leer.GetValue("STATS", "BANCO")) + Subasta.MejorOferta)
        Call LogearEventoDeSubasta("El Oro fue depositado en la boveda Correctamente!.")
        Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "INIT", "MENSAJEINFORMACION", _
                "Subastador te ha dejado un mensaje: ¡Has vendido tu item! Te deposite el oro en el sistema de finanzas Goliath.")
    Else
        UserList(Subastador.ArrayIndex).Stats.GLD = UserList(Subastador.ArrayIndex).Stats.GLD + Subasta.MejorOferta
        Call WriteLocaleMsg(Subastador.ArrayIndex, "1441", e_FontTypeNames.FONTTYPE_SUBASTA, PonerPuntos(Subasta.MejorOferta)) ' Msg1441=Felicitaciones, has ganado ¬1 monedas de oro de tú subasta.
        Call WriteUpdateGold(Subastador.ArrayIndex)
        Call LogearEventoDeSubasta("Oro entregado en la billetera")
    End If
    Call ResetearSubasta
    Exit Sub
FinalizarSubasta_Err:
    Call TraceError(Err.Number, Err.Description, "ModSubasta.FinalizarSubasta", Erl)
End Sub

Public Sub ResetearSubasta()
    On Error GoTo ResetearSubasta_Err
    Subasta.HaySubastaActiva = False
    Subasta.ObjSubastado = 0
    Subasta.ObjSubastadoCantidad = 0
    Subasta.OfertaInicial = 0
    Subasta.Subastador = ""
    Subasta.MejorOferta = 0
    Subasta.Comprador = ""
    Subasta.HuboOferta = False
    Subasta.TiempoRestanteSubasta = 0
    Subasta.MinutosDeSubasta = 0
    Subasta.PosibleCancelo = False
    Call LogearEventoDeSubasta("Subasta finalizada." & data & " a las " & Time)
    Call LogearEventoDeSubasta( _
            "#################################################################################################################################################################################################")
    Exit Sub
ResetearSubasta_Err:
    Call TraceError(Err.Number, Err.Description, "ModSubasta.ResetearSubasta", Erl)
End Sub

Public Sub DevolverItem()
    On Error GoTo DevolverItem_Err
    Dim ObjVendido   As t_Obj
    Dim tUser        As t_UserReference
    Dim Subastador   As t_UserReference
    Dim Leer         As New clsIniManager
    Dim FileUser     As String
    Dim SlotEnBoveda As Integer
    Dim PosMap       As Byte
    Dim PosX         As Byte
    Dim PosY         As Byte
    Call LogearEventoDeSubasta("Subasta cancelada por falta de ofertas, devolviendo items...")
    FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"
    ObjVendido.ObjIndex = Subasta.ObjSubastado
    ObjVendido.amount = Subasta.ObjSubastadoCantidad
    tUser = NameIndex(Subasta.Subastador)
    If Not IsValidUserRef(tUser) Then
        Call LogearEventoDeSubasta("El usuario vendedor de subasta se encuentra offline, intentando depositar en boveda")
        Call Leer.Initialize(FileUser)
        SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1
        If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
            Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
            Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
            Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", _
                    "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te deposite el item en la boveda.")
        Else
            PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
            PosX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
            PosY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))
            If MapData(PosMap, PosX, PosY).ObjInfo.ObjIndex > 0 Then Exit Sub
            If MapData(PosMap, PosX, PosY).TileExit.Map > 0 Then Exit Sub
            If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub
            If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
            Call MakeObj(ObjVendido, PosMap, PosX, PosY)
            Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & PosX & "-" & PosY)
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", _
                    "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")
        End If
    Else
        Subastador = NameIndex(Subasta.Subastador)
        If Not MeterItemEnInventario(Subastador.ArrayIndex, ObjVendido) Then
            Call TirarItemAlPiso(UserList(Subastador.ArrayIndex).pos, ObjVendido)
            Call LogearEventoDeSubasta("Se tiro al piso el item.")
        End If
        Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")
    End If
    Call ResetearSubasta
    Exit Sub
DevolverItem_Err:
    Call TraceError(Err.Number, Err.Description, "ModSubasta.DevolverItem", Erl)
End Sub

Public Sub CancelarSubasta()
    On Error GoTo CancelarSubasta_Err
    Dim ObjVendido   As t_Obj
    Dim tUser        As t_UserReference
    Dim Subastador   As t_UserReference
    Dim Leer         As New clsIniManager
    Dim FileUser     As String
    Dim SlotEnBoveda As Integer
    Dim PosMap       As Byte
    Dim PosX         As Byte
    Dim PosY         As Byte
    Call LogearEventoDeSubasta("Subasta cancelada.")
    FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"
    ObjVendido.ObjIndex = Subasta.ObjSubastado
    ObjVendido.amount = Subasta.ObjSubastadoCantidad
    tUser = NameIndex(Subasta.Subastador)
    If Not IsValidUserRef(tUser) Then
        Call LogearEventoDeSubasta("El usuario de subasta se encuentra offline, intentando depositar en boveda")
        Call Leer.Initialize(FileUser)
        SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1
        If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
            Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
            Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
            Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te deposite el item en la boveda.")
        Else
            PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
            PosX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
            PosY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))
            If MapData(PosMap, PosX, PosY).ObjInfo.ObjIndex > 0 Then Exit Sub
            If MapData(PosMap, PosX, PosY).TileExit.Map > 0 Then Exit Sub
            If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub
            If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
            Call MakeObj(ObjVendido, PosMap, PosX, PosY)
            Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & PosX & "-" & PosY)
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", _
                    "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")
        End If
    Else
        Subastador = NameIndex(Subasta.Subastador)
        If Not MeterItemEnInventario(Subastador.ArrayIndex, ObjVendido) Then
            Call TirarItemAlPiso(UserList(Subastador.ArrayIndex).pos, ObjVendido)
            Call LogearEventoDeSubasta("Se tiro al piso el item.")
        End If
        Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")
        UserList(tUser.ArrayIndex).flags.Subastando = False
    End If
    Call ResetearSubasta
    Exit Sub
CancelarSubasta_Err:
    Call TraceError(Err.Number, Err.Description, "ModSubasta.CancelarSubasta", Erl)
End Sub
