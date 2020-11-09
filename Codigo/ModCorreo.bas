Attribute VB_Name = "ModCorreo"
Option Explicit

Public Sub SortCorreos(ByVal UserIndex As Integer)

    Dim LoopC       As Long

    Dim counter     As Long

    Dim withoutRead As Long

    Dim tempCorreo  As UserCorreo

    Dim indexviejo  As Byte

    Dim i           As Byte

    UserList(UserIndex).Correo.CantCorreo = UserList(UserIndex).Correo.CantCorreo - 1

    For LoopC = 1 To MAX_CORREOS_SLOTS

        If UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = "" Then
            indexviejo = LoopC
        
            For i = indexviejo To MAX_CORREOS_SLOTS - 1
                UserList(UserIndex).Correo.Mensaje(i).Remitente = UserList(UserIndex).Correo.Mensaje(i + 1).Remitente
                UserList(UserIndex).Correo.Mensaje(i).Fecha = UserList(UserIndex).Correo.Mensaje(i + 1).Fecha
                UserList(UserIndex).Correo.Mensaje(i).Item = UserList(UserIndex).Correo.Mensaje(i + 1).Item
                UserList(UserIndex).Correo.Mensaje(i).ItemCount = UserList(UserIndex).Correo.Mensaje(i + 1).ItemCount
                UserList(UserIndex).Correo.Mensaje(i).Mensaje = UserList(UserIndex).Correo.Mensaje(i + 1).Mensaje
                UserList(UserIndex).Correo.Mensaje(i).Leido = UserList(UserIndex).Correo.Mensaje(i + 1).Leido
            Next i

            LoopC = MAX_CORREOS_SLOTS

        End If
    
    Next LoopC

    Call WriteListaCorreo(UserIndex, True)

End Sub

'Note: UserIndex is Emisor, and UserName is Receptor.
Public Function AddCorreo(ByVal UserIndex As Integer, ByRef UserName As String, ByRef message As String, ByVal ObjArray As String, ByVal FinalCount As Byte) As Boolean

    On Error GoTo Errhandler

    Dim ReceptIndex As Integer

    Dim Index       As Byte

    ReceptIndex = NameIndex(UserName)

    If ReceptIndex > 0 And UserList(UserIndex).flags.BattleModo = 0 Then
        Index = SearchIndexFreeCorreo(ReceptIndex)
    
        If Index >= 1 And Index <= MAX_CORREOS_SLOTS Then
            UserList(ReceptIndex).Correo.CantCorreo = UserList(ReceptIndex).Correo.CantCorreo + 1
            UserList(ReceptIndex).Correo.Mensaje(Index).Remitente = UserList(UserIndex).name
            UserList(ReceptIndex).Correo.Mensaje(Index).Mensaje = message
            UserList(ReceptIndex).Correo.Mensaje(Index).Item = ObjArray
            UserList(ReceptIndex).Correo.Mensaje(Index).ItemCount = FinalCount
            UserList(ReceptIndex).Correo.Mensaje(Index).Leido = 0
            UserList(ReceptIndex).Correo.Mensaje(Index).Fecha = Date & " - " & Time
        
            ' UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 75
        
            'If FinalCount <> 0 Then
            ' UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - (1500)
            
            ' End If
            ' Call WriteUpdateUserStats(UserIndex)
        
            '''
            Call WriteConsoleMsg(ReceptIndex, "Has recibido un nuevo mensaje de " & UserList(UserIndex).name & " ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFOIAO)
            UserList(ReceptIndex).Correo.NoLeidos = 1
            Call WriteCorreoPicOn(ReceptIndex)
            ' Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
        
            'El mensaje fue enviado correctamente.
            If UserIndex <> 0 Then Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_INFOIAO)
            
            AddCorreo = True
            Exit Function
        Else

            If UserIndex <> 0 Then Call WriteConsoleMsg(UserIndex, "No hay mas espacio para correos.", FontTypeNames.FONTTYPE_INFOIAO)
            'Logear que no se pudo enviar.
            AddCorreo = False
            Exit Function

        End If

    Else
        'base de datos:

        If PersonajeExiste(UserName) Then
    
            Dim Leer       As New clsIniReader

            Dim FileUser   As String

            Dim CantCorreo As Byte
            
            FileUser = CharPath & UCase$(UserName) & ".chr"
            
            Call Leer.Initialize(FileUser)
            
            CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))

            If CantCorreo = 60 Then
                Call WriteConsoleMsg(UserIndex, "El correo del personaje esta lleno.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Function

            End If

            ' UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 75
    
            ' If FinalCount > 0 Then
            '  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1500
            ' End If
            Call WriteUpdateUserStats(UserIndex)
        
            AddCorreo = GrabarNuevoCorreoInChar(UserName, UserIndex, message, ObjArray, FinalCount)
        
            If UserIndex <> 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("174", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
                Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_INFOIAO) 'El mensaje fue enviado correctamente.
                Call WriteUpdateUserStats(UserIndex)

            End If

            Exit Function
        Else

            If UserIndex <> 0 Then Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFOIAO) 'El personaje no existe o se encuentra baneado.
        
            AddCorreo = False
            Exit Function

        End If

    End If

Errhandler:
    AddCorreo = False

End Function

Public Sub BorrarCorreoMail(ByVal UserIndex As Integer, ByVal Index As Byte)
    UserList(UserIndex).Correo.Mensaje(Index).Remitente = ""
    Call SortCorreos(UserIndex)

End Sub

Public Sub ExtractItemCorreo(ByVal UserIndex As Integer, ByVal Index As Byte)

    If UserList(UserIndex).Correo.Mensaje(Index).ItemCount <= 0 Then Exit Sub
    
    Dim ObjAMeter As obj

    Dim i         As Byte

    Dim rdata     As String

    Dim Item      As String

    Dim ObjIndex  As Long

    Dim Cantidad  As String
    
    For i = 1 To UserList(UserIndex).Correo.Mensaje(Index).ItemCount
    
        rdata = Right$(UserList(UserIndex).Correo.Mensaje(Index).Item, Len(UserList(UserIndex).Correo.Mensaje(Index).Item))
        Item = (ReadField(i, rdata, Asc("@")))
                
        rdata = Left$(Item, Len(Item))
        ObjIndex = (ReadField(1, rdata, Asc("-")))
        
        rdata = Right$(Item, Len(Item))
        Cantidad = (ReadField(2, rdata, Asc("-")))

        ObjAMeter.ObjIndex = ObjIndex
        ObjAMeter.Amount = Cantidad
        
        If Not MeterItemEnInventario(UserIndex, ObjAMeter) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjAMeter)

        End If

    Next i

    UserList(UserIndex).Correo.Mensaje(Index).ItemCount = 0
    UserList(UserIndex).Correo.Mensaje(Index).Item = 0
    Call WriteListaCorreo(UserIndex, True)

End Sub

Public Sub ReadMessageCorreo(ByVal UserIndex As Integer, ByVal Index As Byte)
    UserList(UserIndex).Correo.Mensaje(Index).Leido = 1
    UserList(UserIndex).Correo.MensajesSinLeer = UserList(UserIndex).Correo.MensajesSinLeer - 1

    '   Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
End Sub

Private Function SearchIndexFreeCorreo(ByVal UserIndex As Integer) As Byte

    Dim LoopC As Long

    For LoopC = 1 To MAX_CORREOS_SLOTS

        If UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = "" Then
            SearchIndexFreeCorreo = LoopC
            Exit Function

        End If

    Next LoopC

    SearchIndexFreeCorreo = -1 'No hay más espacio. :P

End Function

Private Function GrabarNuevoCorreoInChar(ByRef UserName As String, ByVal EmisorIndex As Integer, ByRef message As String, ByVal ObjArray As String, ByVal ItemCount As Byte) As Boolean

    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
            
        Dim Leer       As New clsIniReader

        Dim FileUser   As String

        Dim CantCorreo As Byte
            
        FileUser = CharPath & UCase$(UserName) & ".chr"
            
        Call Leer.Initialize(FileUser)
            
        CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))
        CantCorreo = CantCorreo + 1
            
        Call WriteVar(FileUser, "Correo", "CantCorreo", CByte(CantCorreo))
        Call WriteVar(FileUser, "Correo", "REMITENTE" & CantCorreo, UserList(EmisorIndex).name)
        Call WriteVar(FileUser, "Correo", "MENSAJE" & CantCorreo, message)
        Call WriteVar(FileUser, "Correo", "Item" & CantCorreo, ObjArray)
        Call WriteVar(FileUser, "Correo", "ItemCount" & CantCorreo, ItemCount)
        Call WriteVar(FileUser, "Correo", "LEIDO" & CantCorreo, 0)
        Call WriteVar(FileUser, "Correo", "DATE" & CantCorreo, Date)
        Call WriteVar(FileUser, "Correo", "NoLeidos", CByte(1))
        GrabarNuevoCorreoInChar = True
            
        'Call WriteChatOverHead(UserIndex, "¡El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
    Else
        ' Call WriteChatOverHead(EmisorIndex, "El usuario es inexistente.",
        '  Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_SERVER)
            
        GrabarNuevoCorreoInChar = False

    End If

End Function

Private Function GrabarNuevoCorreoInCharBySubasta(ByRef Comprador As String, ByVal Vendedor As String, ByRef message As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer) As Boolean

    If FileExist(CharPath & UCase$(Comprador) & ".chr", vbNormal) Then
            
        Dim Leer       As New clsIniReader

        Dim FileUser   As String

        Dim CantCorreo As Byte
            
        FileUser = CharPath & UCase$(Comprador) & ".chr"
            
        Call Leer.Initialize(FileUser)
            
        CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))
        CantCorreo = CantCorreo + 1
            
        Call WriteVar(FileUser, "Correo", "CantCorreo", CByte(CantCorreo))
        Call WriteVar(FileUser, "Correo", "REMITENTE" & CantCorreo, Vendedor)
        Call WriteVar(FileUser, "Correo", "MENSAJE" & CantCorreo, message)
        Call WriteVar(FileUser, "Correo", "ItemCount" & CantCorreo, 1)
        Call WriteVar(FileUser, "Correo", "Item" & CantCorreo, "@" & ObjIndex & "-" & Cantidad & "@")
        Call WriteVar(FileUser, "Correo", "LEIDO" & CantCorreo, 0)
        Call WriteVar(FileUser, "Correo", "DATE" & CantCorreo, Date)
        Call WriteVar(FileUser, "Correo", "NoLeidos", CByte(1))
        GrabarNuevoCorreoInCharBySubasta = True
            
        'Call WriteChatOverHead(UserIndex, "¡El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
    Else
        ' Call WriteChatOverHead(EmisorIndex, "El usuario es inexistente.",
        '  Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_SERVER)
            
        GrabarNuevoCorreoInCharBySubasta = False

    End If

End Function

Public Function AddCorreoBySubastador(ByVal Vendedor As String, ByRef Comprador As String, ByRef message As String, ByVal obj As Integer, ByVal Cantidad As Integer) As Boolean

    On Error GoTo Errhandler

    Dim ReceptIndex As Integer

    Dim Index       As Byte

    Dim ObjIndex    As Integer

    ReceptIndex = NameIndex(Comprador)

    ObjIndex = obj

    If ReceptIndex > 0 Then
        Index = SearchIndexFreeCorreo(ReceptIndex)
    
        If Index >= 1 And Index <= MAX_CORREOS_SLOTS Then
            UserList(ReceptIndex).Correo.CantCorreo = UserList(ReceptIndex).Correo.CantCorreo + 1
            UserList(ReceptIndex).Correo.Mensaje(Index).Remitente = Vendedor
            UserList(ReceptIndex).Correo.Mensaje(Index).Mensaje = message
            UserList(ReceptIndex).Correo.Mensaje(Index).ItemCount = 1
            UserList(ReceptIndex).Correo.Mensaje(Index).Item = ObjIndex & "-" & Cantidad & "@"
            UserList(ReceptIndex).Correo.Mensaje(Index).Leido = 0
            UserList(ReceptIndex).Correo.Mensaje(Index).Fecha = Date & " - " & Time
        
            '''
            Call WriteConsoleMsg(ReceptIndex, "Has recibido un nuevo mensaje de " & Vendedor & " ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFOIAO)
            UserList(ReceptIndex).Correo.NoLeidos = 1
            Call WriteCorreoPicOn(ReceptIndex)
            ' Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
            
            AddCorreoBySubastador = True
            Exit Function
        Else
        
            AddCorreoBySubastador = False
            Exit Function

        End If

    Else
        'base de datos:

        If PersonajeExiste(Comprador) Then
    
            Dim Leer       As New clsIniReader

            Dim FileUser   As String

            Dim CantCorreo As Byte
            
            FileUser = CharPath & UCase$(Comprador) & ".chr"
            
            Call Leer.Initialize(FileUser)
            
            CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))

            If CantCorreo = 60 Then
                'Call WriteConsoleMsg(UserIndex, "El correo del personaje esta lleno.", FontTypeNames.FONTTYPE_INFOIAO)
                AddCorreoBySubastador = False
                Exit Function

            End If
        
            AddCorreoBySubastador = GrabarNuevoCorreoInCharBySubasta(Comprador, Vendedor, message, ObjIndex, Cantidad)

            Exit Function
        Else
        
            AddCorreoBySubastador = False
            Exit Function

        End If

    End If

Errhandler:
    AddCorreoBySubastador = False

End Function

