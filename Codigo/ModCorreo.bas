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
130     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.SortCorreos", Erl)
132     Resume Next
        
End Sub

'Note: UserIndex is Emisor, and UserName is Receptor.
Public Function AddCorreo(ByVal UserIndex As Integer, ByRef UserName As String, ByRef message As String, ByVal ObjArray As String, ByVal FinalCount As Byte) As Boolean

        On Error GoTo ErrHandler

        Dim ReceptIndex As Integer

        Dim index       As Byte

100     ReceptIndex = NameIndex(UserName)

102     If ReceptIndex > 0 Then
104         index = SearchIndexFreeCorreo(ReceptIndex)
    
106         If index >= 1 And index <= MAX_CORREOS_SLOTS Then
108             UserList(ReceptIndex).Correo.CantCorreo = UserList(ReceptIndex).Correo.CantCorreo + 1
110             UserList(ReceptIndex).Correo.Mensaje(index).Remitente = UserList(UserIndex).name
112             UserList(ReceptIndex).Correo.Mensaje(index).Mensaje = message
114             UserList(ReceptIndex).Correo.Mensaje(index).Item = ObjArray
116             UserList(ReceptIndex).Correo.Mensaje(index).ItemCount = FinalCount
118             UserList(ReceptIndex).Correo.Mensaje(index).Leido = 0
120             UserList(ReceptIndex).Correo.Mensaje(index).Fecha = Date & " - " & Time
        
                ' UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 75
        
                'If FinalCount <> 0 Then
                ' UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - (1500)
            
                ' End If
                ' Call WriteUpdateUserStats(UserIndex)
        
                '''
122             Call WriteConsoleMsg(ReceptIndex, "Has recibido un nuevo mensaje de " & UserList(UserIndex).name & " ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFOIAO)
124             UserList(ReceptIndex).Correo.NoLeidos = 1
126             Call WriteCorreoPicOn(ReceptIndex)
                ' Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
        
                'El mensaje fue enviado correctamente.
128             If UserIndex <> 0 Then Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_INFOIAO)
            
130             AddCorreo = True
                Exit Function
            Else

132             If UserIndex <> 0 Then Call WriteConsoleMsg(UserIndex, "No hay mas espacio para correos.", FontTypeNames.FONTTYPE_INFOIAO)
                'Logear que no se pudo enviar.
134             AddCorreo = False
                Exit Function

            End If

        Else
            'base de datos:

136         If PersonajeExiste(UserName) Then
    
                Dim Leer       As New clsIniManager

                Dim FileUser   As String

                Dim CantCorreo As Byte
            
138             FileUser = CharPath & UCase$(UserName) & ".chr"
            
140             Call Leer.Initialize(FileUser)
            
142             CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))

144             If CantCorreo = 60 Then
146                 Call WriteConsoleMsg(UserIndex, "El correo del personaje esta lleno.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Function

                End If

                ' UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 75
    
                ' If FinalCount > 0 Then
                '  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 1500
                ' End If
148             Call WriteUpdateUserStats(UserIndex)
        
150             AddCorreo = GrabarNuevoCorreoInChar(UserName, UserIndex, message, ObjArray, FinalCount)
        
152             If UserIndex <> 0 Then
154                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("174", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
156                 Call WriteConsoleMsg(UserIndex, "Mensaje enviado.", FontTypeNames.FONTTYPE_INFOIAO) 'El mensaje fue enviado correctamente.
158                 Call WriteUpdateUserStats(UserIndex)

                End If

                Exit Function
            Else

160             If UserIndex <> 0 Then Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_INFOIAO) 'El personaje no existe o se encuentra baneado.
        
162             AddCorreo = False
                Exit Function

            End If

        End If

ErrHandler:
164     AddCorreo = False

End Function

Public Sub BorrarCorreoMail(ByVal UserIndex As Integer, ByVal index As Byte)
        
        On Error GoTo BorrarCorreoMail_Err
        
100     UserList(UserIndex).Correo.Mensaje(index).Remitente = ""
102     Call SortCorreos(UserIndex)

        
        Exit Sub

BorrarCorreoMail_Err:
104     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.BorrarCorreoMail", Erl)
106     Resume Next
        
End Sub

Public Sub ExtractItemCorreo(ByVal UserIndex As Integer, ByVal index As Byte)
        
        On Error GoTo ExtractItemCorreo_Err
        

100     If UserList(UserIndex).Correo.Mensaje(index).ItemCount <= 0 Then Exit Sub
    
        Dim ObjAMeter As obj

        Dim i         As Byte

        Dim rdata     As String

        Dim Item      As String

        Dim ObjIndex  As Long

        Dim Cantidad  As String
    
102     For i = 1 To UserList(UserIndex).Correo.Mensaje(index).ItemCount
    
104         rdata = Right$(UserList(UserIndex).Correo.Mensaje(index).Item, Len(UserList(UserIndex).Correo.Mensaje(index).Item))
106         Item = (ReadField(i, rdata, Asc("@")))
                
108         rdata = Left$(Item, Len(Item))
110         ObjIndex = (ReadField(1, rdata, Asc("-")))
        
112         rdata = Right$(Item, Len(Item))
114         Cantidad = (ReadField(2, rdata, Asc("-")))

116         ObjAMeter.ObjIndex = ObjIndex
118         ObjAMeter.amount = Cantidad
        
120         If Not MeterItemEnInventario(UserIndex, ObjAMeter) Then
122             Call TirarItemAlPiso(UserList(UserIndex).Pos, ObjAMeter)

            End If

124     Next i

126     UserList(UserIndex).Correo.Mensaje(index).ItemCount = 0
128     UserList(UserIndex).Correo.Mensaje(index).Item = 0
130     Call WriteListaCorreo(UserIndex, True)

        
        Exit Sub

ExtractItemCorreo_Err:
132     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.ExtractItemCorreo", Erl)
134     Resume Next
        
End Sub

Public Sub ReadMessageCorreo(ByVal UserIndex As Integer, ByVal index As Byte)
        
        On Error GoTo ReadMessageCorreo_Err
        
100     UserList(UserIndex).Correo.Mensaje(index).Leido = 1
102     UserList(UserIndex).Correo.MensajesSinLeer = UserList(UserIndex).Correo.MensajesSinLeer - 1

        '   Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
        
        Exit Sub

ReadMessageCorreo_Err:
104     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.ReadMessageCorreo", Erl)
106     Resume Next
        
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

108     SearchIndexFreeCorreo = -1 'No hay más espacio. :P

        
        Exit Function

SearchIndexFreeCorreo_Err:
110     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.SearchIndexFreeCorreo", Erl)
112     Resume Next
        
End Function

Private Function GrabarNuevoCorreoInChar(ByRef UserName As String, ByVal EmisorIndex As Integer, ByRef message As String, ByVal ObjArray As String, ByVal ItemCount As Byte) As Boolean
        
        On Error GoTo GrabarNuevoCorreoInChar_Err
        

100     If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
            
            Dim Leer       As New clsIniManager

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
            
            'Call WriteChatOverHead(UserIndex, "¡El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            ' Call WriteChatOverHead(EmisorIndex, "El usuario es inexistente.",
            '  Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_SERVER)
            
128         GrabarNuevoCorreoInChar = False

        End If

        
        Exit Function

GrabarNuevoCorreoInChar_Err:
130     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.GrabarNuevoCorreoInChar", Erl)
132     Resume Next
        
End Function

Private Function GrabarNuevoCorreoInCharBySubasta(ByRef Comprador As String, ByVal Vendedor As String, ByRef message As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer) As Boolean
        
        On Error GoTo GrabarNuevoCorreoInCharBySubasta_Err
        

100     If FileExist(CharPath & UCase$(Comprador) & ".chr", vbNormal) Then
            
            Dim Leer       As New clsIniManager

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
            
            'Call WriteChatOverHead(UserIndex, "¡El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", NpcList(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            ' Call WriteChatOverHead(EmisorIndex, "El usuario es inexistente.",
            '  Call WriteConsoleMsg(UserIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_SERVER)
            
128         GrabarNuevoCorreoInCharBySubasta = False

        End If

        
        Exit Function

GrabarNuevoCorreoInCharBySubasta_Err:
130     Call RegistrarError(Err.Number, Err.Description, "ModCorreo.GrabarNuevoCorreoInCharBySubasta", Erl)
132     Resume Next
        
End Function

Public Function AddCorreoBySubastador(ByVal Vendedor As String, ByRef Comprador As String, ByRef message As String, ByVal obj As Integer, ByVal Cantidad As Integer) As Boolean

        On Error GoTo ErrHandler

        Dim ReceptIndex As Integer

        Dim index       As Byte

        Dim ObjIndex    As Integer

100     ReceptIndex = NameIndex(Comprador)

102     ObjIndex = obj

104     If ReceptIndex > 0 Then
106         index = SearchIndexFreeCorreo(ReceptIndex)
    
108         If index >= 1 And index <= MAX_CORREOS_SLOTS Then
110             UserList(ReceptIndex).Correo.CantCorreo = UserList(ReceptIndex).Correo.CantCorreo + 1
112             UserList(ReceptIndex).Correo.Mensaje(index).Remitente = Vendedor
114             UserList(ReceptIndex).Correo.Mensaje(index).Mensaje = message
116             UserList(ReceptIndex).Correo.Mensaje(index).ItemCount = 1
118             UserList(ReceptIndex).Correo.Mensaje(index).Item = ObjIndex & "-" & Cantidad & "@"
120             UserList(ReceptIndex).Correo.Mensaje(index).Leido = 0
122             UserList(ReceptIndex).Correo.Mensaje(index).Fecha = Date & " - " & Time
        
                '''
124             Call WriteConsoleMsg(ReceptIndex, "Has recibido un nuevo mensaje de " & Vendedor & " ve a un correo local para leerlo.", FontTypeNames.FONTTYPE_INFOIAO)
126             UserList(ReceptIndex).Correo.NoLeidos = 1
128             Call WriteCorreoPicOn(ReceptIndex)
                ' Call WriteCorreoUpdateCount(ReceptIndex, UserList(ReceptIndex).Correo.MensajesSinLeer)
            
130             AddCorreoBySubastador = True
                Exit Function
            Else
        
132             AddCorreoBySubastador = False
                Exit Function

            End If

        Else
            'base de datos:

134         If PersonajeExiste(Comprador) Then
    
                Dim Leer       As New clsIniManager

                Dim FileUser   As String

                Dim CantCorreo As Byte
            
136             FileUser = CharPath & UCase$(Comprador) & ".chr"
            
138             Call Leer.Initialize(FileUser)
            
140             CantCorreo = CByte(Leer.GetValue("Correo", "CantCorreo"))

142             If CantCorreo = 60 Then
                    'Call WriteConsoleMsg(UserIndex, "El correo del personaje esta lleno.", FontTypeNames.FONTTYPE_INFOIAO)
144                 AddCorreoBySubastador = False
                    Exit Function

                End If
        
146             AddCorreoBySubastador = GrabarNuevoCorreoInCharBySubasta(Comprador, Vendedor, message, ObjIndex, Cantidad)

                Exit Function
            Else
        
148             AddCorreoBySubastador = False
                Exit Function

            End If

        End If

ErrHandler:
150     AddCorreoBySubastador = False

End Function

