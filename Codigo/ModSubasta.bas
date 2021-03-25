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
128         Subasta.ObjSubastadoCantidad = MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.Amount
130         Subasta.Subastador = UserList(UserIndex).name
132         UserList(UserIndex).Counters.TiempoParaSubastar = 15
134         Call WriteChatOverHead(UserIndex, "Escribe /OFERTAINICIAL (cantidad) para comenzar la subasta. ¡Tienes 15 segundos!", NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
136         Call EraseObj(Subasta.ObjSubastadoCantidad, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
138         UserList(UserIndex).flags.Subastando = True
            Exit Sub

        End If

        
        Exit Sub

IniciarSubasta_Err:
140     Call RegistrarError(Err.Number, Err.Description, "ModSubasta.IniciarSubasta", Erl)
142     Resume Next
        
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

        Dim Leer         As New clsIniReader

        Dim FileUser     As String

        Dim SlotEnBoveda As Integer

        Dim PosMap       As Byte

        Dim posX         As Byte

        Dim posY         As Byte

100     FileUser = CharPath & UCase$(Subasta.Comprador) & ".chr"

102     ObjVendido.ObjIndex = Subasta.ObjSubastado
104     ObjVendido.Amount = Subasta.ObjSubastadoCantidad
106     tUser = NameIndex(Subasta.Comprador)

114     If tUser <= 0 Then
116         Call LogearEventoDeSubasta("El usuario ganador de subasta se encuentra offline, intentando depositar en boveda")
118         Call Leer.Initialize(FileUser)
120         SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

122         If SlotEnBoveda < MAX_BANCOINVENTORY_SLOTS Then
                
124             Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
126             Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
128             Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te deposite el item en la boveda.")
130             Call LogearEventoDeSubasta("El items fue depositado en la boveda del comprador correctamente.")
            Else
132             Call LogearEventoDeSubasta("Se esta intentando enviar por correo el item.")

134             If AddCorreoBySubastador("Subastador", Subasta.Comprador, "¡Felicitaciones! Ganaste la subasta.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
136                 Call LogearEventoDeSubasta("El items fue enviado al comprador por correo correctamente.")
138                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Te envie el item por correo.")
            
                Else
            
140                 PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
142                 posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
144                 posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

146                 If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

148                 If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

150                 If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

152                 If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
154                 Call MakeObj(ObjVendido, PosMap, posX, posY)
156                 Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
158                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has ganado la subasta! Como no tenias espacio ni en tu boveda ni en el correo, tuve que tirarlo en tu ultima posicion.")

                End If

            End If

        Else

160         If Not MeterItemEnInventario(NameIndex(Subasta.Comprador), ObjVendido) Then
162             Call TirarItemAlPiso(UserList(NameIndex(Subasta.Comprador)).Pos, ObjVendido)

            End If

164         Call LogearEventoDeSubasta("Se entrego el item en mano.")
166         Call WriteConsoleMsg(tUser, "Felicitaciones, has ganado la subasta.", FontTypeNames.FONTTYPE_SUBASTA)

        End If

        Dim Descuento As Long

168     Descuento = Subasta.MejorOferta / 100 * 5
170     Subasta.MejorOferta = Subasta.MejorOferta - Descuento
        
178     If NameIndex(Subasta.Subastador) <= 0 Then
180         Call LogearEventoDeSubasta("El subastador se encontraba offline cuando se le tenia que dar el oro, depositando en el banco.")
182         Call Leer.Initialize(CharPath & UCase$(Subasta.Subastador) & ".chr")
184         Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "STATS", "Banco", CLng(Leer.GetValue("STATS", "BANCO")) + Subasta.MejorOferta)
186         Call LogearEventoDeSubasta("El Oro fue depositado en la boveda Correctamente!.")
188         Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: ¡Has vendido tu item! Te deposite el oro en el sistema de finanzas Goliath.")
        Else
190         UserList(NameIndex(Subasta.Subastador)).Stats.GLD = UserList(NameIndex(Subasta.Subastador)).Stats.GLD + Subasta.MejorOferta
        
192         Call WriteConsoleMsg(NameIndex(Subasta.Subastador), "Felicitaciones, has ganado " & PonerPuntos(Subasta.MejorOferta) & " monedas de oro de tú subasta.", FontTypeNames.FONTTYPE_SUBASTA)
        
194         Call WriteUpdateGold(NameIndex(Subasta.Subastador))
196         Call LogearEventoDeSubasta("Oro entregado en la billetera")
        
        End If

198     Call ResetearSubasta

        
        Exit Sub

FinalizarSubasta_Err:
200     Call RegistrarError(Err.Number, Err.Description, "ModSubasta.FinalizarSubasta", Erl)
202     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.Description, "ModSubasta.LogearEventoDeSubasta", Erl)
110     Resume Next
        
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
122     Call LogearEventoDeSubasta("Subasta finalizada." & Data & " a las " & Time)
124     Call LogearEventoDeSubasta("#################################################################################################################################################################################################")

        
        Exit Sub

ResetearSubasta_Err:
126     Call RegistrarError(Err.Number, Err.Description, "ModSubasta.ResetearSubasta", Erl)
128     Resume Next
        
End Sub

Public Sub DevolverItem()
        
        On Error GoTo DevolverItem_Err
        

        Dim ObjVendido   As obj

        Dim tUser        As Integer

        Dim Leer         As New clsIniReader

        Dim FileUser     As String

        Dim SlotEnBoveda As Integer

        Dim PosMap       As Byte

        Dim posX         As Byte

        Dim posY         As Byte

100     Call LogearEventoDeSubasta("Subasta cancelada por falta de ofertas, devolviendo items...")

102     FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"

104     ObjVendido.ObjIndex = Subasta.ObjSubastado
106     ObjVendido.Amount = Subasta.ObjSubastadoCantidad
108     tUser = NameIndex(Subasta.Subastador)

120     If tUser <= 0 Then
    
122         Call LogearEventoDeSubasta("El usuario vendedor de subasta se encuentra offline, intentando depositar en boveda")
124         Call Leer.Initialize(FileUser)
126         SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

128         If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
                
130             Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
132             Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
134             Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
136             Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te deposite el item en la boveda.")
            Else
            
138             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Su subasta fue cancelada por falta de ofertas", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
                    'Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, te devolvi el item por correo.")
140                 Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se envio por correo el item")
                Else
                
142                 PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
144                 posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
146                 posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

148                 If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

150                 If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

152                 If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

154                 If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
156                 Call MakeObj(ObjVendido, PosMap, posX, posY)
158                 Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
160                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada por falta de ofertas, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")

                End If

            End If

        Else

162         If Not MeterItemEnInventario(NameIndex(Subasta.Subastador), ObjVendido) Then
164             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Tu subasta fue cancelada por falta de ofertas.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
166                 Call LogearEventoDeSubasta("Se envio por correo el item.")
                Else
                
168                 Call TirarItemAlPiso(UserList(NameIndex(Subasta.Subastador)).Pos, ObjVendido)
170                 Call LogearEventoDeSubasta("Se tiro al piso el item.")

                End If

            End If

172         Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")

        End If

174     Call ResetearSubasta

        
        Exit Sub

DevolverItem_Err:
176     Call RegistrarError(Err.Number, Err.Description, "ModSubasta.DevolverItem", Erl)
178     Resume Next
        
End Sub

Public Sub CancelarSubasta()
        
        On Error GoTo CancelarSubasta_Err
        

        Dim ObjVendido   As obj

        Dim tUser        As Integer

        Dim Leer         As New clsIniReader

        Dim FileUser     As String

        Dim SlotEnBoveda As Integer

        Dim PosMap       As Byte

        Dim posX         As Byte

        Dim posY         As Byte

100     Call LogearEventoDeSubasta("Subasta cancelada.")

102     FileUser = CharPath & UCase$(Subasta.Subastador) & ".chr"

104     ObjVendido.ObjIndex = Subasta.ObjSubastado
106     ObjVendido.Amount = Subasta.ObjSubastadoCantidad
108     tUser = NameIndex(Subasta.Subastador)

120     If tUser <= 0 Then
122         Call LogearEventoDeSubasta("El usuario de subasta se encuentra offline, intentando depositar en boveda")
124         Call Leer.Initialize(FileUser)
126         SlotEnBoveda = CInt(Leer.GetValue("BancoInventory", "CantidadItems")) + 1

128         If SlotEnBoveda - 1 < MAX_BANCOINVENTORY_SLOTS Then
                
130             Call WriteVar(FileUser, "BancoInventory", "Obj" & SlotEnBoveda, Subasta.ObjSubastado & "-" & Subasta.ObjSubastadoCantidad)
132             Call WriteVar(FileUser, "BancoInventory", "CantidadItems", SlotEnBoveda)
134             Call LogearEventoDeSubasta("El items fue depositado en la boveda del subastador correctamente.")
136             Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te deposite el item en la boveda.")
            Else
            
138             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Su subasta fue cancelada", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
140                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, te devolvi el item por correo.")
142                 Call LogearEventoDeSubasta("La boveda del usuario estaba llena, se envio por correo el item")
                Else
                
144                 PosMap = CInt(ReadField(1, Leer.GetValue("INIT", "Position"), 45))
146                 posX = CInt(ReadField(2, Leer.GetValue("INIT", "Position"), 45))
148                 posY = CInt(ReadField(3, Leer.GetValue("INIT", "Position"), 45))

150                 If MapData(PosMap, posX, posY).ObjInfo.ObjIndex > 0 Then Exit Sub

152                 If MapData(PosMap, posX, posY).TileExit.Map > 0 Then Exit Sub

154                 If Subasta.ObjSubastado < 1 Or Subasta.ObjSubastado > NumObjDatas Then Exit Sub

156                 If LenB(ObjData(Subasta.ObjSubastado).name) = 0 Then Exit Sub
158                 Call MakeObj(ObjVendido, PosMap, posX, posY)
160                 Call LogearEventoDeSubasta("El correo del usuario estaba lleno, se tiro en la posicion:" & PosMap & "-" & posX & "-" & posY)
162                 Call WriteVar(FileUser, "INIT", "MENSAJEINFORMACION", "Subastador te ha dejado un mensaje: Tu subasta fue cancelada, como no tenias lugar ni en tu correo ni boveda, tuve que tirarlo en tu ultimo posicion.")

                End If

            End If

        Else

164         If Not MeterItemEnInventario(NameIndex(Subasta.Subastador), ObjVendido) Then
166             If AddCorreoBySubastador("Subastador", Subasta.Subastador, "Tu subasta fue cancelada.", Subasta.ObjSubastado, Subasta.ObjSubastadoCantidad) Then
168                 Call LogearEventoDeSubasta("Se envio por correo el item.")
                Else
                
170                 Call TirarItemAlPiso(UserList(NameIndex(Subasta.Subastador)).Pos, ObjVendido)
172                 Call LogearEventoDeSubasta("Se tiro al piso el item.")

                End If

            End If

174         Call LogearEventoDeSubasta("Se entrego el item en mano del subastador.")
176         UserList(tUser).flags.Subastando = False
            
        End If

178     Call ResetearSubasta

        
        Exit Sub

CancelarSubasta_Err:
180     Call RegistrarError(Err.Number, Err.Description, "ModSubasta.CancelarSubasta", Erl)
182     Resume Next
        
End Sub
