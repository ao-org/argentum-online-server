Attribute VB_Name = "ModSubasta"

Public Type tSubastas

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

Public Subasta As tSubastas

Dim Logear     As String

Public Sub IniciarSubasta(UserIndex)

    If UserList(UserIndex).flags.Subastando = True And Not Subasta.HaySubastaActiva Then
        Call WriteChatOverHead(UserIndex, "Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. Te quedan: " & UserList(UserIndex).Counters.TiempoParaSubastar & " segundos... ¡Apurate!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
        Exit Sub

    End If

    If Subasta.HaySubastaActiva = True Then
        Call WriteChatOverHead(UserIndex, "Oye amigo, espera tu turno, estoy subastando en este momento.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
        Exit Sub

    End If

    If Not MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
        Call WriteChatOverHead(UserIndex, "¿Pues Acaso el aire está en venta ahora? ¡Bribón!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
        Exit Sub

    End If
    
    If Not ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Subastable = 1 Then
        Call WriteChatOverHead(UserIndex, "Aquí solo subastamos items que sean valiosos. ¡Largate de acá Bribón!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
        Exit Sub

    End If
    
    If UserList(UserIndex).flags.Subastando = True Then 'Practicamente imposible que pase... pero por si las dudas
        Call WriteChatOverHead(UserIndex, "Tu ya estas subastando! Esto a quedado logeado.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbRed)
        Logear = "El usuario que ya estaba subastando pudo subastar otro item" & Date & " - " & Time
        Call LogearEventoDeSubasta(Logear)
        Exit Sub

    End If
    
    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
        Subasta.ObjSubastado = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex
        Subasta.ObjSubastadoCantidad = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).ObjInfo.Amount
        Subasta.Subastador = UserList(UserIndex).name
        UserList(UserIndex).Counters.TiempoParaSubastar = 15
        Call WriteChatOverHead(UserIndex, "Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. ¡Tienes 15 segundos!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
        Call EraseObj(Subasta.ObjSubastadoCantidad, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
        UserList(UserIndex).flags.Subastando = True
        Exit Sub

    End If

End Sub

Public Sub FinalizarSubasta()
    'Primero Damos el objeto subastado
    'Despues el oro al subastador

    '1) nos fijamos si el usuario que gano la subasta esta online,
    'si esta online le ponemos el objeto en el inventario, si no tiene
    'lugar en el inventario lo tiramos al piso,
    'si esta offline el item se va a depositar en boveda.
    'Si no tiene lugar en boveda, el item se tira en la ultima posicion
    'donde el user estubo parado.

    '2)Nos fijamos si esta online, si esta online, le damos el oro a la billetera,
    'si esta offline se deposita en el banco.
    'El sistema le cobra un 10% del precio de venta, por uso de servicio.

    Dim ObjVendido   As obj

    Dim tUser        As Integer

    Dim Leer         As New clsIniReader

    Dim FileUser     As String

    Dim SlotEnBoveda As Integer

    Dim PosMap       As Byte

    Dim posX         As Byte

    Dim posY         As Byte

    FileUser = CharPath & UCase$(Subasta.Comprador) & ".chr"

    ObjVendido.ObjIndex = Subasta.ObjSubastado
    ObjVendido.Amount = Subasta.ObjSubastadoCantidad
    tUser = NameIndex(Subasta.Comprador)

    Dim EstaBattle As Boolean

    If tUser > 0 Then
        If UserList(tUser).flags.BattleModo = 1 Then
            EstaBattle = True

        End If

    End If

    If tUser <= 0 Or EstaBattle Then
        Call LogearEventoDeSubasta("El usuario ganador de subasta se encuentra offline, intentando depositar en boveda")
        Call Leer.Initialize(FileUser)
        SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

        If SlotEnBoveda < MAX_BANCOINVENTORY_SLOTS Then
                
            Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
            Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te deposite el item en la boveda.")
            Call LogearEventoDeSubasta("El items fue depositado en la boveda del comprador correctamente.")
        Else
            Call LogearEventoDeSubasta("Se esta intentando enviar por correo el item.")

            If AddCorreoBySubastador("Subastador", Subasta.Comprador, "¡Felicitaciones! Ganaste la subasta.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                Call LogearEventoDeSubasta("El items fue enviado al comprador por correo correctamente.")
                Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te envie el item por correo.")
            
            Else
            
                PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
                posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
                posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

                If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

                If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

                If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

                If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
                Call MakeObj(ObjVendido, PosMap, posX, posY)
                Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
                Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Como no tenias espacio ni en tu boveda ni en el correo, tuve que tirarlo en tu ultima posicion.")

            End If

        End If

    Else

        If Not MeterItemEnInventario(NameIndex(Subasta.Comprador), ObjVendido) Then
            Call TirarItemAlPiso(UserList(NameIndex(Subasta.Comprador)).Pos, ObjVendido)

        End If

        Call LogearEventoDeSubasta("Se entrego el item en mano.")
        Call WriteConsoleMsg(tUser, "Felicitaciones, has ganado la subasta.", FontTypeNames.FONTTYPE_SUBASTA)

    End If

    Dim Descuento As Long

    Descuento = Subasta.MejorOferta / 100 * 5
    Subasta.MejorOferta = Subasta.MejorOferta - Descuento
    
    If NameIndex(Subasta.Subastador) > 0 Then
        If UserList(NameIndex(Subasta.Subastador)).flags.BattleModo = 1 Then
            EstaBattle = True

        End If

    End If
    
    If NameIndex(Subasta.Subastador) <= 0 Or EstaBattle Then
        Call LogearEventoDeSubasta("El subastador se encontraba offline cuando se le tenia que dar el oro, depositando en el banco.")
        Call Leer.Initialize(CharPath & UCase$(Subasta.Subastador) & ".chr")
        Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "STATS", "Banco", CLng(Leer.GetValue("STATS", "BANCO")) + Subasta.MejorOferta)
        Call LogearEventoDeSubasta("El Oro fue depositado en la boveda Correctamente!.")
        Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has vendido tu item! Te deposite el oro en el sistema de finanzas Goliath.")
    Else
        UserList(NameIndex(Subasta.Subastador)).Stats.GLD = UserList(NameIndex(Subasta.Subastador)).Stats.GLD + Subasta.MejorOferta
        
        Call WriteConsoleMsg(NameIndex(Subasta.Subastador), "Felicitaciones, has ganado " & Subasta.MejorOferta & " monedas de oro de tú subasta.", FontTypeNames.FONTTYPE_SUBASTA)
        
        Call WriteUpdateGold(NameIndex(Subasta.Subastador))
        Call LogearEventoDeSubasta("Oro entregado en la billetera")
        
    End If

    Call ResetearSubasta

End Sub

Public Sub LogearEventoDeSubasta(Logeo As String)

    Dim n As Integer

    n = FreeFile
    Open App.Path & "\LOGS\subastas.log" For Append Shared As n
    Print #n, Logeo
    Close #n

End Sub

Public Sub ResetearSubasta()
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
    Call LogearEventoDeSubasta("#################################################################################################################################################################################################")

End Sub

Public Sub DevolverItem()

    Dim ObjVendido   As obj

    Dim tUser        As Integer

    Dim Leer         As New clsIniReader

    Dim FileUser     As String

    Dim SlotEnBoveda As Integer

    Dim PosMap       As Byte

    Dim posX         As Byte

    Dim posY         As Byte

    Call LogearEventoDeSubasta("Subasta cancelada por falta de ofertas, devolviendo items...")

    FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"

    ObjVendido.ObjIndex = Subasta.ObjSubastado
    ObjVendido.Amount = Subasta.ObjSubastadoCantidad
    tUser = NameIndex(Subasta.Subastador)

    Dim EstaBattle As Boolean

    If tUser > 0 Then
        If UserList(tUser).flags.BattleModo = 1 Then
            EstaBattle = True
            UserList(tUser).flags.Subastando = False
            UserList(tUser).Counters.TiempoParaSubastar = 0

        End If

    End If

    If tUser <= 0 Or EstaBattle Then
    
        Call LogearEventoDeSubasta("El usuario vendedor de subasta se encuentra offline, intentando depositar en boveda")
        Call Leer.Initialize(FileUser)
        SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

        If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
                
            Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
            Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
            Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te deposite el item en la boveda.")
        Else
            
            If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Su subasta fue cancelada por falta de ofertas", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                'Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te devolvi el item por correo.")
                Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se envio por correo el item")
            Else
                
                PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
                posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
                posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

                If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

                If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

                If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

                If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
                Call MakeObj(ObjVendido, PosMap, posX, posY)
                Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
                Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")

            End If

        End If

    Else

        If Not MeterItemEnInventario(NameIndex(Subasta.Subastador), ObjVendido) Then
            If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Tu subasta fue cancelada por falta de ofertas.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                Call LogearEventoDeSubasta("Se envio por correo el item.")
            Else
                
                Call TirarItemAlPiso(UserList(NameIndex(Subasta.Subastador)).Pos, ObjVendido)
                Call LogearEventoDeSubasta("Se tiro al piso el item.")

            End If

        End If

        Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")

    End If

    Call ResetearSubasta

End Sub

Public Sub CancelarSubasta()

    Dim ObjVendido   As obj

    Dim tUser        As Integer

    Dim Leer         As New clsIniReader

    Dim FileUser     As String

    Dim SlotEnBoveda As Integer

    Dim PosMap       As Byte

    Dim posX         As Byte

    Dim posY         As Byte

    Call LogearEventoDeSubasta("Subasta cancelada.")

    FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"

    ObjVendido.ObjIndex = Subasta.ObjSubastado
    ObjVendido.Amount = Subasta.ObjSubastadoCantidad
    tUser = NameIndex(Subasta.Subastador)

    Dim EstaBattle As Boolean

    If tUser > 0 Then
        If UserList(tUser).flags.BattleModo = 1 Then
            EstaBattle = True
            UserList(tUser).flags.Subastando = False
            UserList(tUser).Counters.TiempoParaSubastar = 0

        End If

    End If

    If tUser <= 0 Or EstaBattle Then
        Call LogearEventoDeSubasta("El usuario de subasta se encuentra offline, intentando depositar en boveda")
        Call Leer.Initialize(FileUser)
        SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

        If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
                
            Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
            Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
            Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
            Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te deposite el item en la boveda.")
        Else
            
            If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Su subasta fue cancelada", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te devolvi el item por correo.")
                Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se envio por correo el item")
            Else
                
                PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
                posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
                posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

                If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

                If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

                If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

                If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
                Call MakeObj(ObjVendido, PosMap, posX, posY)
                Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
                Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")

            End If

        End If

    Else

        If Not MeterItemEnInventario(NameIndex(Subasta.Subastador), ObjVendido) Then
            If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Tu subasta fue cancelada.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                Call LogearEventoDeSubasta("Se envio por correo el item.")
            Else
                
                Call TirarItemAlPiso(UserList(NameIndex(Subasta.Subastador)).Pos, ObjVendido)
                Call LogearEventoDeSubasta("Se tiro al piso el item.")

            End If

        End If

        Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")
        UserList(tUser).flags.Subastando = False
            
    End If

    Call ResetearSubasta

End Sub
