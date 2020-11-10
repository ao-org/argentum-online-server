Attribute VB_Name = "ModCorreo"
Option Explicit

Public Sub SortCorreos(ByVal UserIndex As Integer)
        
        On Error GoTo SortCorreos_Err
        

        Dim LoopC       As Long

        Dim counter     As Long

        Dim withoutRead As Long

        Dim tempCorreo  As UserCorreo

        Dim indexviejo  As Byte

        Dim i           As Byte

100     UserList(UserIndex).Correo.CantCorreo = UserList(UserIndex).Correo.CantCorreo - 1

102     For LoopC = 1 To MAX_CORREOS_SLOTS

104         If UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = "" Then
106             indexviejo = LoopC
        
108             For i = indexviejo To MAX_CORREOS_SLOTS - 1
110                 UserList(UserIndex).Correo.Mensaje(i).Remitente = UserList(UserIndex).Correo.Mensaje(i + 1).Remitente
112                 UserList(UserIndex).Correo.Mensaje(i).Fecha = UserList(UserIndex).Correo.Mensaje(i + 1).Fecha
114                 UserList(UserIndex).Correo.Mensaje(i).Item = UserList(UserIndex).Correo.Mensaje(i + 1).Item
116                 UserList(UserIndex).Correo.Mensaje(i).ItemCount = UserList(UserIndex).Correo.Mensaje(i + 1).ItemCount
118                 UserList(UserIndex).Correo.Mensaje(i).Mensaje = UserList(UserIndex).Correo.Mensaje(i + 1).Mensaje
120                 UserList(UserIndex).Correo.Mensaje(i).Leido = UserList(UserIndex).Correo.Mensaje(i + 1).Leido
122             Next i

124             LoopC = MAX_CORREOS_SLOTS

            End If
    
126     Next LoopC

128     Call WriteListaCorreo(UserIndex, True)

        
        Exit Sub

SortCorreos_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.SortCorreos", Erl)
        Resume Next
        
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
        
        On Error GoTo BorrarCorreoMail_Err
        
100     UserList(UserIndex).Correo.Mensaje(Index).Remitente = ""
102     Call SortCorreos(UserIndex)

        
        Exit Sub

BorrarCorreoMail_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.BorrarCorreoMail", Erl)
        Resume Next
        
End Sub

Public Sub ExtractItemCorreo(ByVal UserIndex As Integer, ByVal Index As Byte)
        
        On Error GoTo ExtractItemCorreo_Err
        

100     If UserList(UserIndex).Correo.Mensaje(Index).ItemCount <= 0 Then Exit Sub
    
        Dim ObjAMeter As obj

        Dim i         As Byte

        Dim rdata     As String

        Dim Item      As String

        Dim ObjIndex  As Long

        Dim Cantidad  As String
    
102     For i = 1 To UserList(UserIndex).Correo.Mensaje(Index).ItemCount
    
104         rdata = Right$(UserList(UserIndex).Correo.Mensaje(Index).Item, Len(UserList(UserIndex).Correo.Mensaje(Index).Item))
106         Item = (ReadField(i, rdata, Asc("@")))
                
108         rdata = Left$(Item, Len(Item))
110         ObjIndex = (ReadField(1, rdata, Asc("-")))
        
112         rdata = Right$(Item, Len(Item))
114         Cantidad = (ReadField(2, rdata, Asc("-")))

116         ObjAMeter.ObjIndex = ObjIndex
118         ObjAMeter.Amount = Cantidad
        
120         If Not MeterItemEnInventario(UserIndex, ObjAMeter) Then
122             Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjAMeter)

            End If

124     Next i

126     UserList(UserIndex).Correo.Mensaje(Index).ItemCount = 0
128     UserList(UserIndex).Correo.Mensaje(Index).Item = 0
130     Call WriteListaCorreo(UserIndex, True)

        
        Exit Sub

ExtractItemCorreo_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.ExtractItemCorreo", Erl)
        Resume Next
        
End Sub

Public Sub ReadMessageCorreo(ByVal UserIndex As Integer, ByVal Index As Byte)
        
        On Error GoTo ReadMessageCorreo_Err
        
100     UserList(UserIndex).Correo.Mensaje(Index).Leido = 1
102     UserList(UserIndex).Correo.MensajesSinLeer = UserList(UserIndex).Correo.MensajesSinLeer - 1

        '   Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
        
        Exit Sub

ReadMessageCorreo_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.ReadMessageCorreo", Erl)
        Resume Next
        
End Sub

Private Function SearchIndexFreeCorreo(ByVal UserIndex As Integer) As Byte
        
        On Error GoTo SearchIndexFreeCorreo_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To MAX_CORREOS_SLOTS

102         If UserList(UserIndex).Correo.Mensaje(LoopC).Remitente = "" Then
104             SearchIndexFreeCorreo = LoopC
                Exit Function

            End If

106     Next LoopC

108     SearchIndexFreeCorreo = -1 'No hay m�s espacio. :P

        
        Exit Function

SearchIndexFreeCorreo_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.SearchIndexFreeCorreo", Erl)
        Resume Next
        
End Function

Private Function GrabarNuevoCorreoInChar(ByRef UserName As String, ByVal EmisorIndex As Integer, ByRef message As String, ByVal ObjArray As String, ByVal ItemCount As Byte) As Boolean
        
        On Error GoTo GrabarNuevoCorreoInChar_Err
        

100     If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
            
            Dim Leer       As New clsIniReader

            Dim FileUser   As String

            Dim CantCorreo As Byte
            
102         FileUser = CharPath & UCase$(UserName) & ".chr"
            
104         Call Leer.Initialize(FileUser)
            
106         CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))
108         CantCorreo = CantCorreo + 1
            
110         Call WriteVar(FileUser, "Correo", "CantCorreo", CByte(CantCorreo))
112         Call WriteVar(FileUser, "Correo", "REMITENTE" & CantCorreo, UserList(EmisorIndex).name)
114         Call WriteVar(FileUser, "Correo", "MENSAJE" & CantCorreo, message)
116         Call WriteVar(FileUser, "Correo", "Item" & CantCorreo, ObjArray)
118         Call WriteVar(FileUser, "Correo", "ItemCount" & CantCorreo, ItemCount)
120         Call WriteVar(FileUser, "Correo", "LEIDO" & CantCorreo, 0)
122         Call WriteVar(FileUser, "Correo", "DATE" & CantCorreo, Date)
124         Call WriteVar(FileUser, "Correo", "NoLeidos", CByte(1))
126         GrabarNuevoCorreoInChar = True
            
            'Call WriteChatOverHead(UserIndex, "�El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            ' Call WriteChatOverHead(EmisorIndex, "El usuario es inexistente.",
            '  Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_SERVER)
            
128         GrabarNuevoCorreoInChar = False

        End If

        
        Exit Function

GrabarNuevoCorreoInChar_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.GrabarNuevoCorreoInChar", Erl)
        Resume Next
        
End Function

Private Function GrabarNuevoCorreoInCharBySubasta(ByRef Comprador As String, ByVal Vendedor As String, ByRef message As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer) As Boolean
        
        On Error GoTo GrabarNuevoCorreoInCharBySubasta_Err
        

100     If FileExist(CharPath & UCase$(Comprador) & ".chr", vbNormal) Then
            
            Dim Leer       As New clsIniReader

            Dim FileUser   As String

            Dim CantCorreo As Byte
            
102         FileUser = CharPath & UCase$(Comprador) & ".chr"
            
104         Call Leer.Initialize(FileUser)
            
106         CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))
108         CantCorreo = CantCorreo + 1
            
110         Call WriteVar(FileUser, "Correo", "CantCorreo", CByte(CantCorreo))
112         Call WriteVar(FileUser, "Correo", "REMITENTE" & CantCorreo, Vendedor)
114         Call WriteVar(FileUser, "Correo", "MENSAJE" & CantCorreo, message)
116         Call WriteVar(FileUser, "Correo", "ItemCount" & CantCorreo, 1)
118         Call WriteVar(FileUser, "Correo", "Item" & CantCorreo, "@" & ObjIndex & "-" & Cantidad & "@")
120         Call WriteVar(FileUser, "Correo", "LEIDO" & CantCorreo, 0)
122         Call WriteVar(FileUser, "Correo", "DATE" & CantCorreo, Date)
124         Call WriteVar(FileUser, "Correo", "NoLeidos", CByte(1))
126         GrabarNuevoCorreoInCharBySubasta = True
            
            'Call WriteChatOverHead(UserIndex, "�El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            ' Call WriteChatOverHead(EmisorIndex, "El usuario es inexistente.",
            '  Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_SERVER)
            
128         GrabarNuevoCorreoInCharBySubasta = False

        End If

        
        Exit Function

GrabarNuevoCorreoInCharBySubasta_Err:
        Call RegistrarError(Err.Number, Err.description, "ModCorreo.GrabarNuevoCorreoInCharBySubasta", Erl)
        Resume Next
        
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

