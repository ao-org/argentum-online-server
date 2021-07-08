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
        
        On Error GoTo IniciarSubasta_Err
        

100     If UserList(UserIndex).flags.Subastando = True And Not Subasta.HaySubastaActiva Then
102         Call WriteChatOverHead(UserIndex, "Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. Te quedan: " & UserList(UserIndex).Counters.TiempoParaSubastar & " segundos... ¡Apurate!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If

104     If Subasta.HaySubastaActiva = True Then
106         Call WriteChatOverHead(UserIndex, "Oye amigo, espera tu turno, estoy subastando en este momento.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If

108     If Not MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
110         Call WriteChatOverHead(UserIndex, "¿Pues Acaso el aire está en venta ahora? ¡Bribón!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If
    
112     If Not ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Subastable = 1 Then
114         Call WriteChatOverHead(UserIndex, "Aquí solo subastamos items que sean valiosos. ¡Largate de acá Bribón!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub

        End If
    
116     If UserList(UserIndex).flags.Subastando = True Then 'Practicamente imposible que pase... pero por si las dudas
118         Call WriteChatOverHead(UserIndex, "Tu ya estas subastando! Esto a quedado logeado.", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbRed)
120         Logear = "El usuario que ya estaba subastando pudo subastar otro item" & Date & " - " & Time
122         Call LogearEventoDeSubasta(Logear)
            Exit Sub

        End If
    
124     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
126         Subasta.ObjSubastado = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex
128         Subasta.ObjSubastadoCantidad = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.amount
130         Subasta.Subastador = UserList(UserIndex).Name
132         UserList(UserIndex).Counters.TiempoParaSubastar = 15
134         Call WriteChatOverHead(UserIndex, "Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. ¡Tienes 15 segundos!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
136         Call EraseObj(Subasta.ObjSubastadoCantidad, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
138         UserList(UserIndex).flags.Subastando = True
            Exit Sub

        End If

        
        Exit Sub

IniciarSubasta_Err:
140     Call TraceError(Err.Number, Err.Description, "ModSubasta.IniciarSubasta", Erl)
142
        
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

        Dim ObjVendido   As obj

        Dim tUser        As Integer

        Dim Leer         As New clsIniManager

        Dim FileUser     As String

        Dim SlotEnBoveda As Integer

        Dim PosMap       As Byte

        Dim posX         As Byte

        Dim posY         As Byte

100     FileUser = CharPath & UCase$(Subasta.Comprador) & ".chr"

102     ObjVendido.ObjIndex = Subasta.ObjSubastado
104     ObjVendido.amount = Subasta.ObjSubastadoCantidad
106     tUser = NameIndex(Subasta.Comprador)

108     If tUser <= 0 Then
110         Call LogearEventoDeSubasta("El usuario ganador de subasta se encuentra offline, intentando depositar en boveda")
112         Call Leer.Initialize(FileUser)
114         SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

116         If SlotEnBoveda < MAX_BANCOINVENTORY_SLOTS Then
                
118             Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
120             Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
122             Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te deposite el item en la boveda.")
124             Call LogearEventoDeSubasta("El items fue depositado en la boveda del comprador correctamente.")
            Else
126             Call LogearEventoDeSubasta("Se esta intentando enviar por correo el item.")

128             If AddCorreoBySubastador("Subastador", Subasta.Comprador, "¡Felicitaciones! Ganaste la subasta.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
130                 Call LogearEventoDeSubasta("El items fue enviado al comprador por correo correctamente.")
132                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te envie el item por correo.")
            
                Else
            
134                 PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
136                 posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
138                 posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

140                 If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

142                 If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

144                 If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

146                 If LenB(ObjData(Subasta.ObjSubastado).Name) = 0 Then Exit Sub
148                 Call MakeObj(ObjVendido, PosMap, posX, posY)
150                 Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
152                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Como no tenias espacio ni en tu boveda ni en el correo, tuve que tirarlo en tu ultima posicion.")

                End If

            End If

        Else

154         If Not MeterItemEnInventario(NameIndex(Subasta.Comprador), ObjVendido) Then
156             Call TirarItemAlPiso(UserList(NameIndex(Subasta.Comprador)).Pos, ObjVendido)

            End If

158         Call LogearEventoDeSubasta("Se entrego el item en mano.")
160         Call WriteConsoleMsg(tUser, "Felicitaciones, has ganado la subasta.", FontTypeNames.FONTTYPE_SUBASTA)

        End If

        Dim Descuento As Long

162     Descuento = Subasta.MejorOferta / 100 * 5
164     Subasta.MejorOferta = Subasta.MejorOferta - Descuento
        
166     If NameIndex(Subasta.Subastador) <= 0 Then
168         Call LogearEventoDeSubasta("El subastador se encontraba offline cuando se le tenia que dar el oro, depositando en el banco.")
170         Call Leer.Initialize(CharPath & UCase$(Subasta.Subastador) & ".chr")
172         Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "STATS", "Banco", CLng(Leer.GetValue("STATS", "BANCO")) + Subasta.MejorOferta)
174         Call LogearEventoDeSubasta("El Oro fue depositado en la boveda Correctamente!.")
176         Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has vendido tu item! Te deposite el oro en el sistema de finanzas Goliath.")
        Else
178         UserList(NameIndex(Subasta.Subastador)).Stats.GLD = UserList(NameIndex(Subasta.Subastador)).Stats.GLD + Subasta.MejorOferta
        
180         Call WriteConsoleMsg(NameIndex(Subasta.Subastador), "Felicitaciones, has ganado " & PonerPuntos(Subasta.MejorOferta) & " monedas de oro de tú subasta.", FontTypeNames.FONTTYPE_SUBASTA)
        
182         Call WriteUpdateGold(NameIndex(Subasta.Subastador))
184         Call LogearEventoDeSubasta("Oro entregado en la billetera")
        
        End If

186     Call ResetearSubasta

        
        Exit Sub

FinalizarSubasta_Err:
188     Call TraceError(Err.Number, Err.Description, "ModSubasta.FinalizarSubasta", Erl)
190
        
End Sub

Public Sub LogearEventoDeSubasta(Logeo As String)
        
        On Error GoTo LogearEventoDeSubasta_Err
        

        Dim n As Integer

100     n = FreeFile
102     Open App.Path & "\LOGS\subastas.log" For Append Shared As n
104     Print #n, Logeo
106     Close #n

        
        Exit Sub

LogearEventoDeSubasta_Err:
108     Call TraceError(Err.Number, Err.Description, "ModSubasta.LogearEventoDeSubasta", Erl)
110
        
End Sub

Public Sub ResetearSubasta()
        
        On Error GoTo ResetearSubasta_Err
        
100     Subasta.HaySubastaActiva = False
102     Subasta.ObjSubastado = 0
104     Subasta.ObjSubastadoCantidad = 0
106     Subasta.OfertaInicial = 0
108     Subasta.Subastador = ""
110     Subasta.MejorOferta = 0
112     Subasta.Comprador = ""
114     Subasta.HuboOferta = False
116     Subasta.TiempoRestanteSubasta = 0
118     Subasta.MinutosDeSubasta = 0
120     Subasta.PosibleCancelo = False
122     Call LogearEventoDeSubasta("Subasta finalizada." & data & " a las " & Time)
124     Call LogearEventoDeSubasta("#################################################################################################################################################################################################")

        
        Exit Sub

ResetearSubasta_Err:
126     Call TraceError(Err.Number, Err.Description, "ModSubasta.ResetearSubasta", Erl)
128
        
End Sub

Public Sub DevolverItem()
        
        On Error GoTo DevolverItem_Err
        

        Dim ObjVendido   As obj

        Dim tUser        As Integer

        Dim Leer         As New clsIniManager

        Dim FileUser     As String

        Dim SlotEnBoveda As Integer

        Dim PosMap       As Byte

        Dim posX         As Byte

        Dim posY         As Byte

100     Call LogearEventoDeSubasta("Subasta cancelada por falta de ofertas, devolviendo items...")

102     FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"

104     ObjVendido.ObjIndex = Subasta.ObjSubastado
106     ObjVendido.amount = Subasta.ObjSubastadoCantidad
108     tUser = NameIndex(Subasta.Subastador)

110     If tUser <= 0 Then
    
112         Call LogearEventoDeSubasta("El usuario vendedor de subasta se encuentra offline, intentando depositar en boveda")
114         Call Leer.Initialize(FileUser)
116         SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

118         If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
                
120             Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
122             Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
124             Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
126             Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te deposite el item en la boveda.")
            Else
            
128             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Su subasta fue cancelada por falta de ofertas", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                    'Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te devolvi el item por correo.")
130                 Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se envio por correo el item")
                Else
                
132                 PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
134                 posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
136                 posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

138                 If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

140                 If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

142                 If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

144                 If LenB(ObjData(Subasta.ObjSubastado).Name) = 0 Then Exit Sub
146                 Call MakeObj(ObjVendido, PosMap, posX, posY)
148                 Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
150                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")

                End If

            End If

        Else

152         If Not MeterItemEnInventario(NameIndex(Subasta.Subastador), ObjVendido) Then
154             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Tu subasta fue cancelada por falta de ofertas.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
156                 Call LogearEventoDeSubasta("Se envio por correo el item.")
                Else
                
158                 Call TirarItemAlPiso(UserList(NameIndex(Subasta.Subastador)).Pos, ObjVendido)
160                 Call LogearEventoDeSubasta("Se tiro al piso el item.")

                End If

            End If

162         Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")

        End If

164     Call ResetearSubasta

        
        Exit Sub

DevolverItem_Err:
166     Call TraceError(Err.Number, Err.Description, "ModSubasta.DevolverItem", Erl)
168
        
End Sub

Public Sub CancelarSubasta()
        
        On Error GoTo CancelarSubasta_Err
        

        Dim ObjVendido   As obj

        Dim tUser        As Integer

        Dim Leer         As New clsIniManager

        Dim FileUser     As String

        Dim SlotEnBoveda As Integer

        Dim PosMap       As Byte

        Dim posX         As Byte

        Dim posY         As Byte

100     Call LogearEventoDeSubasta("Subasta cancelada.")

102     FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"

104     ObjVendido.ObjIndex = Subasta.ObjSubastado
106     ObjVendido.amount = Subasta.ObjSubastadoCantidad
108     tUser = NameIndex(Subasta.Subastador)

110     If tUser <= 0 Then
112         Call LogearEventoDeSubasta("El usuario de subasta se encuentra offline, intentando depositar en boveda")
114         Call Leer.Initialize(FileUser)
116         SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

118         If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
                
120             Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
122             Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
124             Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
126             Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te deposite el item en la boveda.")
            Else
            
128             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Su subasta fue cancelada", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
130                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te devolvi el item por correo.")
132                 Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se envio por correo el item")
                Else
                
134                 PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
136                 posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
138                 posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

140                 If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

142                 If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

144                 If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

146                 If LenB(ObjData(Subasta.ObjSubastado).Name) = 0 Then Exit Sub
148                 Call MakeObj(ObjVendido, PosMap, posX, posY)
150                 Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
152                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")

                End If

            End If

        Else

154         If Not MeterItemEnInventario(NameIndex(Subasta.Subastador), ObjVendido) Then
156             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Tu subasta fue cancelada.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
158                 Call LogearEventoDeSubasta("Se envio por correo el item.")
                Else
                
160                 Call TirarItemAlPiso(UserList(NameIndex(Subasta.Subastador)).Pos, ObjVendido)
162                 Call LogearEventoDeSubasta("Se tiro al piso el item.")

                End If

            End If

164         Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")
166         UserList(tUser).flags.Subastando = False
            
        End If

168     Call ResetearSubasta

        
        Exit Sub

CancelarSubasta_Err:
170     Call TraceError(Err.Number, Err.Description, "ModSubasta.CancelarSubasta", Erl)
172
        
End Sub
