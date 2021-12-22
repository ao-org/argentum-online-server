Attribute VB_Name = "Database"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018
'Rewrited for Argentum20 by Alexis Caraballo (WyroX)
'October 2020

Option Explicit

Public Database_Enabled     As Boolean
Public Database_Driver      As String
Public Database_Source      As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Private Const MAX_ASYNC     As Byte = 20
Private Current_async       As Byte

Private Connection          As ADODB.Connection
Private Connection_async(1 To MAX_ASYNC)    As ADODB.Connection

Private Builder             As cStringBuilder

Public Sub Database_Connect_Async()
        On Error GoTo Database_Connect_AsyncErr
        
        Dim ConnectionID As String

        If Len(Database_Source) <> 0 Then
104         ConnectionID = "DATA SOURCE=" & Database_Source & ";"
        Else
106         ConnectionID = "DRIVER={SQLite3 ODBC Driver};" & "DATABASE=" & App.Path & "/Database.db"
        End If
                
        Dim i As Byte
        
        For i = 1 To MAX_ASYNC
            Set Connection_async(i) = New ADODB.Connection
110         Connection_async(i).CursorLocation = adUseClient
            Connection_async(i).ConnectionString = ConnectionID
112         Call Connection_async(i).Open(, , , adAsyncConnect)
        Next i

        Current_async = 1
        
113     Set Builder = New cStringBuilder
        Call InitDatabase
        
        Exit Sub
    
Database_Connect_AsyncErr:
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - Database_Connect_Async")
End Sub
Public Sub Database_Connect()
        On Error GoTo Database_Connect_Err
        
        Dim ConnectionID As String

        If Len(Database_Source) <> 0 Then
104         ConnectionID = "DATA SOURCE=" & Database_Source & ";"
        Else
106         ConnectionID = "DRIVER={SQLite3 ODBC Driver};" & "DATABASE=" & App.Path & "/Database.db"
        End If
                
        Set Connection = New ADODB.Connection
110     Connection.CursorLocation = adUseClient
        Connection.ConnectionString = ConnectionID

113     Set Builder = New cStringBuilder
        
112     Call Connection.Open

        Call InitDatabase
        
        Exit Sub
    
Database_Connect_Err:
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - Database_Connect")
End Sub

Public Sub Database_Close()
        On Error GoTo Database_Close_Err

        If Connection.State <> adStateClosed Then
            Call Connection.Close
        End If

        Set Connection = Nothing
        
        Exit Sub
     
Database_Close_Err:
        Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.Description)
End Sub

Public Function Query(ByVal Text As String, ParamArray Arguments() As Variant) As ADODB.Recordset
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    
    Command.ActiveConnection = Connection
    Command.CommandText = Text
    Command.CommandType = adCmdText
    Command.Prepared = True
    
    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument

    On Error GoTo Query_Err

    ' Statistics
    If frmMain.chkLogDbPerfomance.Value = 1 Then
        Call GetElapsedTime
    End If
    
    Set Query = Command.Execute()
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.Value = 1 Then
        Call LogPerformance("Query: " & Text & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If

    Exit Function
    
Query_Err:
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - " & vbCrLf & Text)
End Function

Public Function Execute(ByVal Text As String, ParamArray Arguments() As Variant) As Boolean
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    
    Command.ActiveConnection = Connection_async(Current_async)
    Command.CommandText = Text
    Command.CommandType = adCmdText
    Command.Prepared = True

    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument
    
On Error GoTo Execute_Err
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.Value = 1 Then
        Call GetElapsedTime
    End If
    
    Call Command.Execute(, , adAsyncExecute)  ' @TODO: We want some operation to be async
    
    Current_async = Current_async + 1
    
    If Current_async = MAX_ASYNC Then
        Current_async = 1
    End If
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.Value = 1 Then
        Call LogPerformance("Execute: " & Text & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If
    
Execute_Err:
    
    If (Err.Number <> 0) Then
        Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - " & vbCrLf & Text)
    End If
    
    Execute = (Err.Number = 0)
End Function

Public Function Invoke(ByVal Procedure As String, ParamArray Arguments() As Variant) As ADODB.Recordset
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    Dim Affected As Long
    
    Command.ActiveConnection = Connection
    Command.CommandText = Procedure
    Command.CommandType = adCmdStoredProc
    Command.Prepared = True

    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument
    
On Error GoTo Execute_Err
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.Value = 1 Then
        Call GetElapsedTime
    End If
    
    Set Invoke = Command.Execute()
    
    If (Not Invoke Is Nothing And Invoke.EOF) Then
        Set Invoke = Nothing
    End If
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.Value = 1 Then
        Call LogPerformance("Invoke: " & Procedure & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If
    
Execute_Err:
    If (Err.Number <> 0) Then
        Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description)
    End If
End Function

Private Function CreateParameter(ByVal Value As Variant, ByVal Direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter
    Set CreateParameter = New ADODB.Parameter
    
    CreateParameter.Direction = Direction
    
    Select Case VarType(Value)
        Case VbVarType.vbString
            CreateParameter.Type = adBSTR
            CreateParameter.Size = Len(Value)
            CreateParameter.Value = CStr(Value)
        Case VbVarType.vbDecimal
            CreateParameter.Type = adInteger
            CreateParameter.Value = CLng(Value)
        Case VbVarType.vbByte:
            CreateParameter.Type = adTinyInt
            CreateParameter.Value = CByte(Value)
        Case VbVarType.vbInteger
            CreateParameter.Type = adSmallInt
            CreateParameter.Value = CInt(Value)
        Case VbVarType.vbLong
            CreateParameter.Type = adInteger
            CreateParameter.Value = CLng(Value)
        Case VbVarType.vbBoolean
            CreateParameter.Type = adBoolean
            CreateParameter.Value = CBool(Value)
        Case VbVarType.vbSingle
            CreateParameter.Type = adSingle
            CreateParameter.Value = CSng(Value)
        Case VbVarType.vbDouble
            CreateParameter.Type = adDouble
            CreateParameter.Value = CDbl(Value)
    End Select
End Function

' ------------------------------------------------------- VIEJO (necesita Refactor) -------------------------------------------------------------

Public Sub InitDatabase()
    'Reinicio los users online
    Call SetUsersLoggedDatabase(0)

    'Leo el record de usuarios
    RecordUsuarios = LeerRecordUsuariosDatabase()
End Sub

Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

        On Error GoTo ErrorHandler
    
        Dim LoopC As Long
        Dim ParamC As Integer
        Dim Params() As Variant
    
102     With UserList(UserIndex)
        
            Dim i As Integer
104         ReDim Params(0 To 44)

            '  ************ Basic user data *******************
106         Params(PostInc(i)) = .Name
108         Params(PostInc(i)) = .AccountID
110         Params(PostInc(i)) = .Stats.ELV
112         Params(PostInc(i)) = .Stats.Exp
114         Params(PostInc(i)) = .genero
116         Params(PostInc(i)) = .raza
118         Params(PostInc(i)) = .clase
120         Params(PostInc(i)) = .Hogar
122         Params(PostInc(i)) = .Desc
124         Params(PostInc(i)) = .Stats.GLD
126         Params(PostInc(i)) = .Stats.SkillPts
128         Params(PostInc(i)) = .Pos.Map
130         Params(PostInc(i)) = .Pos.X
132         Params(PostInc(i)) = .Pos.Y
134         Params(PostInc(i)) = .Char.Body
136         Params(PostInc(i)) = .Char.Head
138         Params(PostInc(i)) = .Char.WeaponAnim
140         Params(PostInc(i)) = .Char.CascoAnim
142         Params(PostInc(i)) = .Char.ShieldAnim
146         Params(PostInc(i)) = .Invent.ArmourEqpSlot
148         Params(PostInc(i)) = .Invent.WeaponEqpSlot
150         Params(PostInc(i)) = .Invent.EscudoEqpSlot
152         Params(PostInc(i)) = .Invent.CascoEqpSlot
154         Params(PostInc(i)) = .Invent.MunicionEqpSlot
156         Params(PostInc(i)) = .Invent.DañoMagicoEqpSlot
158         Params(PostInc(i)) = .Invent.ResistenciaEqpSlot
160         Params(PostInc(i)) = .Invent.HerramientaEqpSlot
162         Params(PostInc(i)) = .Invent.MagicoSlot
164         Params(PostInc(i)) = .Invent.NudilloSlot
166         Params(PostInc(i)) = .Invent.BarcoSlot
168         Params(PostInc(i)) = .Invent.MonturaSlot
170         Params(PostInc(i)) = .Stats.MinHp
172         Params(PostInc(i)) = .Stats.MaxHp
174         Params(PostInc(i)) = .Stats.MinMAN
176         Params(PostInc(i)) = .Stats.MaxMAN
178         Params(PostInc(i)) = .Stats.MinSta
180         Params(PostInc(i)) = .Stats.MaxSta
182         Params(PostInc(i)) = .Stats.MinHam
184         Params(PostInc(i)) = .Stats.MaxHam
186         Params(PostInc(i)) = .Stats.MinAGU
188         Params(PostInc(i)) = .Stats.MaxAGU
190         Params(PostInc(i)) = .Stats.MinHIT
192         Params(PostInc(i)) = .Stats.MaxHit
194         Params(PostInc(i)) = .flags.Desnudo
196         Params(PostInc(i)) = .Faccion.Status
        
198         Call Query(QUERY_SAVE_MAINPJ, Params)

            ' Para recibir el ID del user
            Dim RS As ADODB.Recordset
            Set RS = Query("SELECT last_insert_rowid()")

202         If RS Is Nothing Then
204             .ID = 1
            Else
206             .ID = val(RS.Fields(0).Value)
            End If
        
            ' ******************* ATRIBUTOS *******************
208         ReDim Params(1 To NUMATRIBUTOS)
210         ParamC = 1
            
212         For LoopC = 1 To NUMATRIBUTOS
214             Params(PostInc(ParamC)) = .Stats.UserAtributos(LoopC)
222         Next LoopC
        
            Call Query(QUERY_INSERT_ATTRIBUTES, .ID, Params)
        
            ' ******************* SPELLS **********************
226         ReDim Params(MAXUSERHECHIZOS * 3 - 1)
228         ParamC = 0
        
230         For LoopC = 1 To MAXUSERHECHIZOS
232             Params(ParamC) = .ID
234             Params(ParamC + 1) = LoopC
236             Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
238             ParamC = ParamC + 3
240         Next LoopC

            Call Query(QUERY_SAVE_SPELLS, Params)
        
            ' ******************* INVENTORY *******************
244         ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
246         ParamC = 0
        
248         For LoopC = 1 To MAX_INVENTORY_SLOTS
250             Params(ParamC) = .ID
252             Params(ParamC + 1) = LoopC
254             Params(ParamC + 2) = .Invent.Object(LoopC).ObjIndex
256             Params(ParamC + 3) = .Invent.Object(LoopC).amount
258             Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
            
260             ParamC = ParamC + 5
262         Next LoopC
        
            Call Query(QUERY_SAVE_INVENTORY, Params)
        
            ' ******************* SKILLS *******************
266         ReDim Params(NUMSKILLS * 3 - 1)
268         ParamC = 0
        
270         For LoopC = 1 To NUMSKILLS
272             Params(ParamC) = .ID
274             Params(ParamC + 1) = LoopC
276             Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
278             ParamC = ParamC + 3
280         Next LoopC
        
            Call Query(QUERY_SAVE_SKILLS, Params)
        
            ' ******************* QUESTS *******************
284         ReDim Params(MAXUSERQUESTS * 2 - 1)
286         ParamC = 0
        
288         For LoopC = 1 To MAXUSERQUESTS
290             Params(ParamC) = .ID
292             Params(ParamC + 1) = LoopC
            
294             ParamC = ParamC + 2
296         Next LoopC
        
            Call Query(QUERY_SAVE_QUESTS, Params)
        
            ' ******************* PETS ********************
300         ReDim Params(MAXMASCOTAS * 3 - 1)
302         ParamC = 0
        
304         For LoopC = 1 To MAXMASCOTAS
306             Params(ParamC) = .ID
308             Params(ParamC + 1) = LoopC
310             Params(ParamC + 2) = 0
            
312             ParamC = ParamC + 3
314         Next LoopC
    
            Call Query(QUERY_SAVE_PETS, Params)
    
        End With

        Exit Sub

ErrorHandler:
    
322     Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserDatabase(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

        On Error GoTo ErrorHandler
    
        Dim Params() As Variant
        Dim LoopC As Long
        Dim ParamC As Long
    
100     Call Builder.Clear

        'Basic user data
102     With UserList(UserIndex)
            
            
            
            
            
104         ReDim Params(89)

            Dim i As Integer
        
106         Params(PostInc(i)) = .Name
108         Params(PostInc(i)) = .Stats.ELV
110         Params(PostInc(i)) = .Stats.Exp
112         Params(PostInc(i)) = .genero
114         Params(PostInc(i)) = .raza
116         Params(PostInc(i)) = .clase
118         Params(PostInc(i)) = .Hogar
120         Params(PostInc(i)) = .Desc
122         Params(PostInc(i)) = .Stats.GLD
124         Params(PostInc(i)) = .Stats.Banco
126         Params(PostInc(i)) = .Stats.SkillPts
128         Params(PostInc(i)) = .flags.MascotasGuardadas
130         Params(PostInc(i)) = .Pos.Map
132         Params(PostInc(i)) = .Pos.X
134         Params(PostInc(i)) = .Pos.Y
136         Params(PostInc(i)) = .MENSAJEINFORMACION
138         Params(PostInc(i)) = .Char.Body
140         Params(PostInc(i)) = .OrigChar.Head
142         Params(PostInc(i)) = .Char.WeaponAnim
144         Params(PostInc(i)) = .Char.CascoAnim
146         Params(PostInc(i)) = .Char.ShieldAnim
148         Params(PostInc(i)) = .Char.Heading
152         Params(PostInc(i)) = .Invent.ArmourEqpSlot
154         Params(PostInc(i)) = .Invent.WeaponEqpSlot
156         Params(PostInc(i)) = .Invent.EscudoEqpSlot
158         Params(PostInc(i)) = .Invent.CascoEqpSlot
160         Params(PostInc(i)) = .Invent.MunicionEqpSlot
162         Params(PostInc(i)) = .Invent.DañoMagicoEqpSlot
164         Params(PostInc(i)) = .Invent.ResistenciaEqpSlot
166         Params(PostInc(i)) = .Invent.HerramientaEqpSlot
168         Params(PostInc(i)) = .Invent.MagicoSlot
170         Params(PostInc(i)) = .Invent.NudilloSlot
172         Params(PostInc(i)) = .Invent.BarcoSlot
174         Params(PostInc(i)) = .Invent.MonturaSlot
176         Params(PostInc(i)) = .Stats.MinHp
178         Params(PostInc(i)) = .Stats.MaxHp
180         Params(PostInc(i)) = .Stats.MinMAN
182         Params(PostInc(i)) = .Stats.MaxMAN
184         Params(PostInc(i)) = .Stats.MinSta
186         Params(PostInc(i)) = .Stats.MaxSta
188         Params(PostInc(i)) = .Stats.MinHam
190         Params(PostInc(i)) = .Stats.MaxHam
192         Params(PostInc(i)) = .Stats.MinAGU
194         Params(PostInc(i)) = .Stats.MaxAGU
196         Params(PostInc(i)) = .Stats.MinHIT
198         Params(PostInc(i)) = .Stats.MaxHit
200         Params(PostInc(i)) = .Stats.NPCsMuertos
202         Params(PostInc(i)) = .Stats.UsuariosMatados
203         Params(PostInc(i)) = .Stats.PuntosPesca
204         Params(PostInc(i)) = .Stats.InventLevel
206         Params(PostInc(i)) = .Stats.ELO
208         Params(PostInc(i)) = .flags.Desnudo
210         Params(PostInc(i)) = .flags.Envenenado
212         Params(PostInc(i)) = .flags.Incinerado
218         Params(PostInc(i)) = .flags.Muerto
220         Params(PostInc(i)) = .flags.Navegando
222         Params(PostInc(i)) = .flags.Paralizado
224         Params(PostInc(i)) = .flags.Montado
226         Params(PostInc(i)) = .flags.Silenciado
228         Params(PostInc(i)) = .flags.MinutosRestantes
230         Params(PostInc(i)) = .flags.SegundosPasados
232         Params(PostInc(i)) = .flags.Pareja
234         Params(PostInc(i)) = .Counters.Pena
236         Params(PostInc(i)) = .flags.VecesQueMoriste
238         Params(PostInc(i)) = (.flags.Privilegios And e_PlayerType.RoyalCouncil)
240         Params(PostInc(i)) = (.flags.Privilegios And e_PlayerType.ChaosCouncil)
242         Params(PostInc(i)) = .Faccion.ArmadaReal
244         Params(PostInc(i)) = .Faccion.FuerzasCaos
246         Params(PostInc(i)) = .Faccion.ciudadanosMatados
248         Params(PostInc(i)) = .Faccion.CriminalesMatados
250         Params(PostInc(i)) = .Faccion.RecibioArmaduraReal
252         Params(PostInc(i)) = .Faccion.RecibioArmaduraCaos
254         Params(PostInc(i)) = .Faccion.RecibioExpInicialReal
256         Params(PostInc(i)) = .Faccion.RecibioExpInicialCaos
258         Params(PostInc(i)) = .Faccion.RecompensasReal
260         Params(PostInc(i)) = .Faccion.RecompensasCaos
262         Params(PostInc(i)) = .Faccion.Reenlistadas
264         Params(PostInc(i)) = .Faccion.NivelIngreso
266         Params(PostInc(i)) = .Faccion.MatadosIngreso
268         Params(PostInc(i)) = .Faccion.NextRecompensa
270         Params(PostInc(i)) = .Faccion.Status
272         Params(PostInc(i)) = .GuildIndex
274         Params(PostInc(i)) = .ChatCombate
276         Params(PostInc(i)) = .ChatGlobal
280         Params(PostInc(i)) = .Stats.Advertencias
282         Params(PostInc(i)) = .flags.ReturnPos.Map
284         Params(PostInc(i)) = .flags.ReturnPos.X
286         Params(PostInc(i)) = .flags.ReturnPos.Y
287         Params(PostInc(i)) = GetTickCount

            ' WHERE block
288         Params(PostInc(i)) = .ID
            
            Call Execute(QUERY_UPDATE_MAINPJ, Params)

            ' ************************** User spells *********************************
334             ReDim Params(MAXUSERHECHIZOS * 3 - 1)
336             ParamC = 0
            
338             For LoopC = 1 To MAXUSERHECHIZOS
340                 Params(ParamC) = .ID
342                 Params(ParamC + 1) = LoopC
344                 Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
                
346                 ParamC = ParamC + 3
348             Next LoopC
                
                Call Execute(QUERY_UPSERT_SPELLS, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modificaron los hechizos. Guardando..."
                .flags.ModificoHechizos = False
            
            ' ************************** User inventory *********************************
366             ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
368             ParamC = 0
            
370             For LoopC = 1 To MAX_INVENTORY_SLOTS
372                 Params(ParamC) = .ID
374                 Params(ParamC + 1) = LoopC
376                 Params(ParamC + 2) = .Invent.Object(LoopC).ObjIndex
378                 Params(ParamC + 3) = .Invent.Object(LoopC).amount
380                 Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
                
382                 ParamC = ParamC + 5
384             Next LoopC

                Call Execute(QUERY_UPSERT_INVENTORY, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico el inventario. Guardando..."
                .flags.ModificoInventario = False
            
            ' ************************** User bank inventory *********************************
402             ReDim Params(MAX_BANCOINVENTORY_SLOTS * 4 - 1)
404             ParamC = 0
            
406             For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
408                 Params(ParamC) = .ID
410                 Params(ParamC + 1) = LoopC
412                 Params(ParamC + 2) = .BancoInvent.Object(LoopC).ObjIndex
414                 Params(ParamC + 3) = .BancoInvent.Object(LoopC).amount
                
416                 ParamC = ParamC + 4
418             Next LoopC
    
                Call Execute(QUERY_SAVE_BANCOINV, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico el inventario del banco. Guardando..."
                .flags.ModificoInventarioBanco = False

            ' ************************** User skills *********************************
436             ReDim Params(NUMSKILLS * 3 - 1)
438             ParamC = 0
            
440             For LoopC = 1 To NUMSKILLS
442                 Params(ParamC) = .ID
444                 Params(ParamC + 1) = LoopC
446                 Params(ParamC + 2) = .Stats.UserSkills(LoopC)
                
448                 ParamC = ParamC + 3
450             Next LoopC
        
                Call Execute(QUERY_UPSERT_SKILLS, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico las habilidades. Guardando..."
                .flags.ModificoSkills = False

            ' ************************** User pets *********************************
468             ReDim Params(MAXMASCOTAS * 3 - 1)
470             ParamC = 0
                Dim petType As Integer
    
472             For LoopC = 1 To MAXMASCOTAS
474                 Params(ParamC) = .ID
476                 Params(ParamC + 1) = LoopC
    
                    'CHOTS | I got this logic from SaveUserToCharfile
478                 If .MascotasIndex(LoopC) > 0 Then
                
480                     If NpcList(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
482                         petType = .MascotasType(LoopC)
                        Else
484                         petType = 0
                        End If
    
                    Else
486                     petType = .MascotasType(LoopC)
    
                    End If
    
488                 Params(ParamC + 2) = petType
                
490                 ParamC = ParamC + 3
492             Next LoopC
                
                Call Execute(QUERY_UPSERT_PETS, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico las mascotas. Guardando..."
                .flags.ModificoMascotas = False

            ' ************************** User quests *********************************
526             Builder.Append "REPLACE INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
            
                Dim Tmp As Integer, LoopK As Long
    
528             For LoopC = 1 To MAXUSERQUESTS
530                 Builder.Append "("
532                 Builder.Append .ID & ", "
534                 Builder.Append LoopC & ", "
536                 Builder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
                
538                 If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
540                     Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs
    
542                     If Tmp Then
    
544                         For LoopK = 1 To Tmp
546                             Builder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                            
548                             If LoopK < Tmp Then
550                                 Builder.Append "-"
                                End If
    
552                         Next LoopK
                        
    
                        End If
    
                    End If
                
554                 Builder.Append "', '"
                
556                 If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                    
558                     Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                        
560                     For LoopK = 1 To Tmp
    
562                         Builder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                        
564                         If LoopK < Tmp Then
566                             Builder.Append "-"
                            End If
                    
568                     Next LoopK
                
                    End If
                
570                 Builder.Append "')"
    
572                 If LoopC < MAXUSERQUESTS Then
574                     Builder.Append ", "
                    End If
    
576             Next LoopC

                Call Execute(Builder.ToString())

584             Call Builder.Clear
                
                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico las quests. Guardando..."
                .flags.ModificoQuests = False
        
            ' ************************** User completed quests *********************************
586         If .QuestStats.NumQuestsDone > 0 Then
                
                
                    ' Armamos la query con los placeholders
590                 Builder.Append "REPLACE INTO quest_done (user_id, quest_id) VALUES "
                
592                 For LoopC = 1 To .QuestStats.NumQuestsDone
594                     Builder.Append "(?, ?)"
                
596                     If LoopC < .QuestStats.NumQuestsDone Then
598                         Builder.Append ", "
                        End If
                
600                 Next LoopC

                    ' Metemos los parametros
604                 ReDim Params(.QuestStats.NumQuestsDone * 2 - 1)
606                 ParamC = 0
                
608                 For LoopC = 1 To .QuestStats.NumQuestsDone
610                     Params(ParamC) = .ID
612                     Params(ParamC + 1) = .QuestStats.QuestsDone(LoopC)
                    
614                     ParamC = ParamC + 2
616                 Next LoopC
        
                    Call Execute(Builder.ToString(), Params)

626                 Call Builder.Clear
                    
                    ' Reseteamos el flag para no volver a guardar.
                    Debug.Print "Se modifico las quests hechas. Guardando..."
                    .flags.ModificoQuestsHechas = False
            End If
            
        End With
    
        Exit Sub

ErrorHandler:
636     Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Sub LoadUserDatabase(ByVal UserIndex As Integer)
        Dim counter As Long
            
        On Error GoTo ErrorHandler

        'Basic user data
100     With UserList(UserIndex)
            
            Dim RS As ADODB.Recordset
            Set RS = Query(QUERY_LOAD_MAINPJ, .Name)

104         If RS Is Nothing Then Exit Sub
            
            If (CLng(RS!account_id) <> UserList(UserIndex).AccountID) Then
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            If (RS!is_banned) Then
                Dim BanNick     As String
                Dim BaneoMotivo As String
                BanNick = RS!banned_by
                BaneoMotivo = RS!ban_reason
                
                If LenB(BanNick) = 0 Then BanNick = "*Error en la base de datos*"
                If LenB(BaneoMotivo) = 0 Then BaneoMotivo = "*No se registra el motivo del baneo.*"
            
                Call WriteShowMessageBox(UserIndex, "Se te ha prohibido la entrada al juego debido a " & BaneoMotivo & ". Esta decisión fue tomada por " & BanNick & ".")
            
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            Dim last_logout As Long
            
            last_logout = val(RS!last_logout)
            
            'Start setting data
106         .ID = RS!ID
108         .Name = RS!Name
110         .Stats.ELV = RS!level
112         .Stats.Exp = RS!Exp
114         .genero = RS!genre_id
116         .raza = RS!race_id
118         .clase = RS!class_id
120         .Hogar = RS!home_id
122         .Desc = RS!Description
124         .Stats.GLD = RS!gold
126         .Stats.Banco = RS!bank_gold
128         .Stats.SkillPts = RS!free_skillpoints
130         .Pos.Map = RS!pos_map
132         .Pos.X = RS!pos_x
134         .Pos.Y = RS!pos_y
136         .MENSAJEINFORMACION = RS!message_info
138         .OrigChar.Body = RS!body_id
140         .OrigChar.Head = RS!head_id
142         .OrigChar.WeaponAnim = RS!weapon_id
144         .OrigChar.CascoAnim = RS!helmet_id
146         .OrigChar.ShieldAnim = RS!shield_id
148         .OrigChar.Heading = RS!Heading
152         .Invent.ArmourEqpSlot = SanitizeNullValue(RS!slot_armour, 0)
154         .Invent.WeaponEqpSlot = SanitizeNullValue(RS!slot_weapon, 0)
156         .Invent.CascoEqpSlot = SanitizeNullValue(RS!slot_helmet, 0)
158         .Invent.EscudoEqpSlot = SanitizeNullValue(RS!slot_shield, 0)
160         .Invent.MunicionEqpSlot = SanitizeNullValue(RS!slot_ammo, 0)
162         .Invent.BarcoSlot = SanitizeNullValue(RS!slot_ship, 0)
164         .Invent.MonturaSlot = SanitizeNullValue(RS!slot_mount, 0)
166         .Invent.DañoMagicoEqpSlot = SanitizeNullValue(RS!slot_dm, 0)
168         .Invent.ResistenciaEqpSlot = SanitizeNullValue(RS!slot_rm, 0)
170         .Invent.NudilloSlot = SanitizeNullValue(RS!slot_knuckles, 0)
172         .Invent.HerramientaEqpSlot = SanitizeNullValue(RS!slot_tool, 0)
174         .Invent.MagicoSlot = SanitizeNullValue(RS!slot_magic, 0)
176         .Stats.MinHp = RS!min_hp
178         .Stats.MaxHp = RS!max_hp
180         .Stats.MinMAN = RS!min_man
182         .Stats.MaxMAN = RS!max_man
184         .Stats.MinSta = RS!min_sta
186         .Stats.MaxSta = RS!max_sta
188         .Stats.MinHam = RS!min_ham
190         .Stats.MaxHam = RS!max_ham
192         .Stats.MinAGU = RS!min_sed
194         .Stats.MaxAGU = RS!max_sed
196         .Stats.MinHIT = RS!min_hit
198         .Stats.MaxHit = RS!max_hit
200         .Stats.NPCsMuertos = RS!killed_npcs
202         .Stats.UsuariosMatados = RS!killed_users
203         .Stats.PuntosPesca = rs!puntos_pesca
204         .Stats.InventLevel = RS!invent_level
206         .Stats.ELO = RS!ELO
208         .flags.Desnudo = RS!is_naked
210         .flags.Envenenado = RS!is_poisoned
211         .flags.Incinerado = RS!is_incinerated
212         .flags.Escondido = False
218         .flags.Ban = RS!is_banned
220         .flags.Muerto = RS!is_dead
222         .flags.Navegando = RS!is_sailing
224         .flags.Paralizado = RS!is_paralyzed
226         .flags.VecesQueMoriste = RS!deaths
228         .flags.Montado = RS!is_mounted
230         .flags.Pareja = RS!spouse
232         .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
234         .flags.Silenciado = RS!is_silenced
236         .flags.MinutosRestantes = RS!silence_minutes_left
238         .flags.SegundosPasados = RS!silence_elapsed_seconds
240         .flags.MascotasGuardadas = RS!pets_saved

246         .flags.ReturnPos.Map = RS!return_map
248         .flags.ReturnPos.X = RS!return_x
250         .flags.ReturnPos.Y = RS!return_y
        
252         .Counters.Pena = RS!counter_pena
        
254         .ChatGlobal = RS!chat_global
256         .ChatCombate = RS!chat_combate

258         If RS!pertenece_consejo_real Then
260             .flags.Privilegios = .flags.Privilegios Or e_PlayerType.RoyalCouncil

            End If

262         If RS!pertenece_consejo_caos Then
264             .flags.Privilegios = .flags.Privilegios Or e_PlayerType.ChaosCouncil

            End If

266         .Faccion.ArmadaReal = RS!pertenece_real
268         .Faccion.FuerzasCaos = RS!pertenece_caos
270         .Faccion.ciudadanosMatados = RS!ciudadanos_matados
272         .Faccion.CriminalesMatados = RS!criminales_matados
274         .Faccion.RecibioArmaduraReal = RS!recibio_armadura_real
276         .Faccion.RecibioArmaduraCaos = RS!recibio_armadura_caos
278         .Faccion.RecibioExpInicialReal = RS!recibio_exp_real
280         .Faccion.RecibioExpInicialCaos = RS!recibio_exp_caos
282         .Faccion.RecompensasReal = RS!recompensas_real
284         .Faccion.RecompensasCaos = RS!recompensas_caos
286         .Faccion.Reenlistadas = RS!Reenlistadas
288         .Faccion.NivelIngreso = SanitizeNullValue(RS!nivel_ingreso, 0)
290         .Faccion.MatadosIngreso = SanitizeNullValue(RS!matados_ingreso, 0)
292         .Faccion.NextRecompensa = SanitizeNullValue(RS!siguiente_recompensa, 0)
294         .Faccion.Status = RS!Status

296         .GuildIndex = SanitizeNullValue(RS!Guild_Index, 0)
            .LastGuildRejection = SanitizeNullValue(RS!guild_rejected_because, vbNullString)
 
298         .Stats.Advertencias = RS!warnings
        
            'User attributes
            Set RS = Query("SELECT * FROM attribute WHERE user_id = ?;", .ID)
    
302         If Not RS Is Nothing Then
                .Stats.UserAtributos(e_Atributos.Fuerza) = RS!strength
                .Stats.UserAtributos(e_Atributos.Agilidad) = RS!agility
                .Stats.UserAtributos(e_Atributos.Constitucion) = RS!constitution
                .Stats.UserAtributos(e_Atributos.Inteligencia) = RS!intelligence
                .Stats.UserAtributos(e_Atributos.Carisma) = RS!charisma

                .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
                .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
                .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
                .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
                .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
            End If
            
            'User spells
            Set RS = Query("SELECT number, spell_id FROM spell WHERE user_id = ?;", .ID)

316         If Not RS Is Nothing Then

320             While Not RS.EOF

322                 .Stats.UserHechizos(RS!Number) = RS!spell_id

324                 RS.MoveNext
                Wend

            End If

            'User pets
            Set RS = Query("SELECT number, pet_id FROM pet WHERE user_id = ?;", .ID)

328         If Not RS Is Nothing Then

332             While Not RS.EOF

334                 .MascotasType(RS!Number) = RS!pet_id
                
336                 If val(RS!pet_id) <> 0 Then
338                     .NroMascotas = .NroMascotas + 1

                    End If

340                 RS.MoveNext
                Wend

            End If

            'User inventory
            Set RS = Query("SELECT number, item_id, is_equipped, amount FROM inventory_item WHERE user_id = ?;", .ID)

            counter = 0
            
344         If Not RS Is Nothing Then

348             While Not RS.EOF

350                 With .Invent.Object(RS!Number)
352                     .ObjIndex = RS!item_id
                
354                     If .ObjIndex <> 0 Then
356                         If LenB(ObjData(.ObjIndex).Name) Then
                                counter = counter + 1
                                
358                             .amount = RS!amount
360                             .Equipped = RS!is_equipped
                            Else
362                             .ObjIndex = 0

                            End If

                        End If

                    End With

364                 RS.MoveNext
                Wend
                
                .Invent.NroItems = counter
            End If

            'User bank inventory
            Set RS = Query("SELECT number, item_id, amount FROM bank_item WHERE user_id = ?;", .ID)
            
            counter = 0
            
368         If Not RS Is Nothing Then

372             While Not RS.EOF

374                 With .BancoInvent.Object(RS!Number)
376                     .ObjIndex = RS!item_id
                
378                     If .ObjIndex <> 0 Then
380                         If LenB(ObjData(.ObjIndex).Name) Then
                                counter = counter + 1
                                
382                             .amount = RS!amount
                            Else
384                             .ObjIndex = 0

                            End If

                        End If

                    End With

386                 RS.MoveNext
                Wend
                
                .BancoInvent.NroItems = counter
            End If
            
            'User skills
            Set RS = Query("SELECT number, value FROM skillpoint WHERE user_id = ?;", .ID)

390         If Not RS Is Nothing Then

394             While Not RS.EOF

396                 .Stats.UserSkills(RS!Number) = RS!Value

398                 RS.MoveNext
                Wend

            End If

            Dim LoopC As Byte
        
            'User quests
            Set RS = Query("SELECT number, quest_id, npcs, npcstarget FROM quest WHERE user_id = ?;", .ID)

402         If Not RS Is Nothing Then

406             While Not RS.EOF

408                 .QuestStats.Quests(RS!Number).QuestIndex = RS!quest_id
                
410                 If .QuestStats.Quests(RS!Number).QuestIndex > 0 Then
412                     If QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs Then

                            Dim NPCs() As String

414                         NPCs = Split(RS!NPCs, "-")
416                         ReDim .QuestStats.Quests(RS!Number).NPCsKilled(1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs)

418                         For LoopC = 1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs
420                             .QuestStats.Quests(RS!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
422                         Next LoopC

                        End If
                    
424                     If QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs Then

                            Dim NPCsTarget() As String

426                         NPCsTarget = Split(RS!NPCsTarget, "-")
428                         ReDim .QuestStats.Quests(RS!Number).NPCsTarget(1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs)

430                         For LoopC = 1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs
432                             .QuestStats.Quests(RS!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
434                         Next LoopC

                        End If

                    End If

436                 RS.MoveNext
                Wend

            End If
        
            'User quests done
            Set RS = Query("SELECT quest_id FROM quest_done WHERE user_id = ?;", .ID)

440         If Not RS Is Nothing Then
442             .QuestStats.NumQuestsDone = RS.RecordCount
                
                If (.QuestStats.NumQuestsDone > 0) Then
444                 ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
    
448                 LoopC = 1
    
450                 While Not RS.EOF
                
452                     .QuestStats.QuestsDone(LoopC) = RS!quest_id
454                     LoopC = LoopC + 1
    
456                     RS.MoveNext
                    Wend
                End If
            End If

            ' Llaves
            Set RS = Query("SELECT key_obj FROM house_key WHERE account_id = ?", .AccountID)

460         If Not RS Is Nothing Then
464             LoopC = 1

466             While Not RS.EOF

468                 .Keys(LoopC) = RS!key_obj
470                 LoopC = LoopC + 1

472                 RS.MoveNext
                Wend

            End If
            UpdateDBIpsValues UserIndex

        End With
        

        Exit Sub

ErrorHandler:
474     Call LogDatabaseError("Error en LoadUserDatabase: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)

End Sub
Public Function UpdateDBIpsValues(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim ipStr As String
        'ipStr = GetDBValue("account", "last_ip", "id", .AccountID)
        
        
100     Dim RS As ADODB.Recordset
        Set RS = Query("SELECT last_ip FROM account WHERE id = ?", .AccountID)

        'Revisamos si recibio un resultado
102     If RS Is Nothing Then Exit Function
        If RS.BOF Or RS.EOF Then Exit Function
        
        'Obtenemos la variable
104     ipStr = RS.Fields(0).Value
        Dim count As Long
        Dim i As Long
        For i = 1 To Len(ipStr)
            If mid(ipStr, i, 1) = ";" Then
                count = count + 1
            End If
        Next i
        
        'Si ya tengo alguna ip guardada
        If count > 0 And count < 5 Then
            
            ReDim ip_list(0 To (count - 1)) As String
            count = count + 1
            ReDim ip_list_new(0 To (count - 1)) As String
            
            ip_list = Split(ipStr, ";")
            
            For i = 0 To (count - 1)
                If .IP = ip_list(i) Then Exit Function
            Next i
            
            For i = 0 To (count - 1)
                ip_list_new(i) = ip_list(i)
            Next i
            
            ip_list_new(count - 1) = .IP
            
        ElseIf count >= 5 Then
        
            ReDim ip_list(0 To (count - 1)) As String
            ReDim ip_list_new(0 To (count - 1)) As String
            
            ip_list = Split(ipStr, ";")
            
            For i = 0 To (count - 1)
                If .IP = ip_list(i) Then Exit Function
            Next i
            
            For i = 1 To (count - 1)
                ip_list_new(i - 1) = ip_list(i)
            Next i
            
            ip_list_new(count - 1) = .IP
            
        Else
            Call Execute("update account set last_ip = ? where id = ?", .IP & ";", .AccountID)
            Exit Function
        End If
        
    
        ipStr = ""
        For i = 0 To (count - 1)
            ipStr = ipStr & ip_list_new(i) & ";"
        Next i
        
        Debug.Print ipStr
        
         Call Execute("update account set last_ip = ? where id = ?", ipStr, .AccountID)
        
    End With
End Function

Public Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
        On Error GoTo ErrorHandler
    
100     Dim RS As ADODB.Recordset
        Set rs = Query("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE LOWER(" & ColumnaTest & ") = ?;", ValueTest)

        'Revisamos si recibio un resultado
102     If RS Is Nothing Then Exit Function
        If RS.BOF Or RS.EOF Then Exit Function
        
        'Obtenemos la variable
104     GetDBValue = RS.Fields(ColumnaGet).Value

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.Description)
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)
        
        Exit Function

GetUserValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserValue", Erl)
End Function

Public Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
        On Error GoTo ErrorHandler

        Call Execute("UPDATE " & Tabla & " SET " & ColumnaSet & " = ? WHERE " & ColumnaTest & " = ?;", ValueSet, ValueTest)

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.Description)
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, Value, "name", CharName)

        Exit Sub

SetUserValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValue", Erl)
End Sub

Private Sub SetUserValueByID(ByVal ID As Long, Columna As String, Value As Variant)
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, Value, "id", ID)

        Exit Sub

SetUserValueByID_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValueByID", Erl)
End Sub

Public Function BANCheckDatabase(Name As String) As Boolean
        
        On Error GoTo BANCheckDatabase_Err
        
100     BANCheckDatabase = CBool(GetUserValue(LCase$(Name), "is_banned"))
  
        Exit Function

BANCheckDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.BANCheckDatabase", Erl)
End Function

Public Function GetUserStatusDatabase(Name As String) As Integer
        
        On Error GoTo GetUserStatusDatabase_Err
        
100     GetUserStatusDatabase = GetUserValue(LCase$(Name), "status")

        
        Exit Function

GetUserStatusDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserStatusDatabase", Erl)

        
End Function

Public Function GetAccountIDDatabase(Name As String) As Long
        
        On Error GoTo GetAccountIDDatabase_Err
        

        Dim Temp As Variant

100     Temp = GetUserValue(LCase$(Name), "account_id")
    
102     If VBA.IsEmpty(Temp) Then
104         GetAccountIDDatabase = -1
        Else
106         GetAccountIDDatabase = val(Temp)

        End If

        
        Exit Function

GetAccountIDDatabase_Err:
108     Call TraceError(Err.Number, Err.Description, "modDatabase.GetAccountIDDatabase", Erl)

        
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountID As Long) As Byte

        On Error GoTo ErrorHandler
    
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT COUNT(*) FROM user WHERE account_id = ?", AccountID)
    
102     If RS Is Nothing Then Exit Function
    
104     GetPersonajesCountByIDDatabase = RS.Fields(0).Value
    
        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountID As Long, Personaje() As t_PersonajeCuenta) As Byte
        
        On Error GoTo GetPersonajesCuentaDatabase_Err
        
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing FROM user WHERE account_id = ?;", AccountID)

102     If RS Is Nothing Then Exit Function
    
104     GetPersonajesCuentaDatabase = RS.RecordCount

        Dim i As Integer
        If GetPersonajesCuentaDatabase = 0 Then Exit Function
108     For i = 1 To GetPersonajesCuentaDatabase
110         Personaje(i).nombre = RS!Name
112         Personaje(i).Cabeza = RS!head_id
114         Personaje(i).clase = RS!class_id
116         Personaje(i).cuerpo = RS!body_id
118         Personaje(i).Mapa = RS!pos_map
120         Personaje(i).posX = RS!pos_x
122         Personaje(i).posY = RS!pos_y
124         Personaje(i).nivel = RS!level
126         Personaje(i).Status = RS!Status
128         Personaje(i).Casco = RS!helmet_id
130         Personaje(i).Escudo = RS!shield_id
132         Personaje(i).Arma = RS!weapon_id
134         Personaje(i).ClanIndex = RS!Guild_Index
        
136         If EsRolesMaster(Personaje(i).nombre) Then
138             Personaje(i).Status = 3
140         ElseIf EsConsejero(Personaje(i).nombre) Then
142             Personaje(i).Status = 4
144         ElseIf EsSemiDios(Personaje(i).nombre) Then
146             Personaje(i).Status = 5
148         ElseIf EsDios(Personaje(i).nombre) Then
150             Personaje(i).Status = 6
152         ElseIf EsAdmin(Personaje(i).nombre) Then
154             Personaje(i).Status = 7

            End If

156         If val(RS!is_dead) = 1 Or val(RS!is_sailing) = 1 Then
158             Personaje(i).Cabeza = 0
            End If
        
160         RS.MoveNext
        Next

        Exit Function

GetPersonajesCuentaDatabase_Err:
162     Call TraceError(Err.Number, Err.Description, "modDatabase.GetPersonajesCuentaDatabase", Erl)

        
End Function

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
        
        On Error GoTo SetUsersLoggedDatabase_Err
        
        Call Query("UPDATE statistics SET value = ? WHERE name = 'online';", NumUsers)
        
        Exit Sub

SetUsersLoggedDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUsersLoggedDatabase", Erl)

        
End Sub

Public Function LeerRecordUsuariosDatabase() As Long
        
        On Error GoTo LeerRecordUsuariosDatabase_Err
        
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT value FROM statistics WHERE name = 'record';")

102     If RS Is Nothing Then Exit Function

104     LeerRecordUsuariosDatabase = val(RS!Value)

        Exit Function

LeerRecordUsuariosDatabase_Err:
106     Call TraceError(Err.Number, Err.Description, "modDatabase.LeerRecordUsuariosDatabase", Erl)

        
End Function

Public Sub SetRecordUsersDatabase(ByVal Record As Long)
        
        On Error GoTo SetRecordUsersDatabase_Err
                
        Call Execute("UPDATE statistics SET value = ? WHERE name = 'record';", CStr(Record))
        
        Exit Sub

SetRecordUsersDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetRecordUsersDatabase", Erl)

        
End Sub

Public Sub SaveVotoDatabase(ByVal ID As Long, ByVal Encuestas As Integer)
        
        On Error GoTo SaveVotoDatabase_Err
        
100     Call SetUserValueByID(ID, "votes_amount", Encuestas)

        
        Exit Sub

SaveVotoDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveVotoDatabase", Erl)

        
End Sub

Public Sub SaveUserBodyDatabase(UserName As String, ByVal Body As Integer)
        
        On Error GoTo SaveUserBodyDatabase_Err
        
100     Call SetUserValue(UserName, "body_id", Body)

        
        Exit Sub

SaveUserBodyDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserBodyDatabase", Erl)

        
End Sub

Public Sub SaveUserHeadDatabase(UserName As String, ByVal Head As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "head_id", Head)

        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)

        
End Sub

Public Sub SaveUserSkillDatabase(UserName As String, ByVal Skill As Integer, ByVal Value As Integer)
        
        On Error GoTo SaveUserSkillDatabase_Err
        
        Call Execute("UPDATE skillpoints SET value = ? WHERE number = ? AND user_id = (SELECT id FROM user WHERE UPPER(name) = ?)", Value, Skill, UCase$(UserName))
        
        Exit Sub

SaveUserSkillDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserSkillDatabase", Erl)

        
End Sub

Public Sub SaveUserSkillsLibres(UserName As String, ByVal SkillsLibres As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "free_skillpoints", SkillsLibres)
        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)

        
End Sub

Public Sub BorrarUsuarioDatabase(Name As String)

        On Error GoTo ErrorHandler

        Call Execute("insert into user_deleted select * from user where name = ?;", Name)
        Call Execute("delete from user where name = ?;", Name)
        Call Execute("UPDATE user_deleted set deleted = CURRENT_TIMESTAMP where name = ?;", Name)

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en BorrarUsuarioDatabase borrando user de la Mysql Database: " & Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanDatabase(UserName As String, Reason As String, BannedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE UPPER(name) = ?;", BannedBy, Reason, UCase$(UserName))
        
102     Call SavePenaDatabase(UserName, "Baneado por: " & BannedBy & " debido a " & Reason)

        Exit Sub

ErrorHandler:
104     Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveWarnDatabase(UserName As String, Reason As String, WarnedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET warnings = warnings + 1 WHERE UPPER(name) = ?;", UCase$(UserName))
        
102     Call SavePenaDatabase(UserName, "Advertencia de: " & WarnedBy & " debido a " & Reason)
    
    Exit Sub

ErrorHandler:
104     Call LogDatabaseError("Error in SaveWarnDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SavePenaDatabase(UserName As String, Reason As String)

        On Error GoTo ErrorHandler

        Dim Query As String
100     Query = "INSERT INTO punishment(user_id, NUMBER, reason)"
102     Query = Query & " SELECT u.id, COUNT(p.number) + 1, ? FROM user u LEFT JOIN punishment p ON p.user_id = u.id WHERE UPPER(u.name) = ?;"
        
        Call Execute(Query, Reason, UCase$(UserName))

        Exit Sub

ErrorHandler:
106     Call LogDatabaseError("Error in SavePenaDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SilenciarUserDatabase(UserName As String, ByVal Tiempo As Integer)
    
        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET is_silenced = 1, silence_minutes_left = ?, silence_elapsed_seconds = 0 WHERE UPPER(name) = ?;", Tiempo, UCase$(UserName))
        
        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SilenciarUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Sub DesilenciarUserDatabase(UserName As String)

        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "is_silenced", 0)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in DesilenciarUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Sub UnBanDatabase(UserName As String)

        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET is_banned = FALSE, banned_by = '', ban_reason = '' WHERE UPPER(name) = ?;", UCase$(UserName))
        
        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanCuentaDatabase(ByVal AccountID As Long, Reason As String, BannedBy As String)

        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE account SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE id = ?;", BannedBy, Reason, AccountID)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SaveBanCuentaDatabase: AccountId=" & AccountID & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub EcharConsejoDatabase(UserName As String)
        
        On Error GoTo EcharConsejoDatabase_Err
        
        Call Execute("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE UPPER(name) = ?;", UCase$(UserName))
        
        Exit Sub

EcharConsejoDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharConsejoDatabase", Erl)

        
End Sub

Public Sub EcharLegionDatabase(UserName As String)
        
        On Error GoTo EcharLegionDatabase_Err
        
        Call Execute("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?;", UCase$(UserName))
        
        Exit Sub

EcharLegionDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharLegionDatabase", Erl)

        
End Sub

Public Sub EcharArmadaDatabase(UserName As String)
        
        On Error GoTo EcharArmadaDatabase_Err
        
        Call Execute("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?;", UCase$(UserName))

        Exit Sub

EcharArmadaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharArmadaDatabase", Erl)

        
End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
        
        On Error GoTo CambiarPenaDatabase_Err
        
        Call Execute("UPDATE punishment SET reason = ? WHERE number = ? AND user_id = (SELECT id from user WHERE UPPER(name) = ?);", Pena, Numero, UCase$(UserName))
        
        Exit Sub

CambiarPenaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.CambiarPenaDatabase", Erl)

        
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT COUNT(*) as punishments FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", UCase$(UserName))

102     If RS Is Nothing Then Exit Function

104     GetUserAmountOfPunishmentsDatabase = RS!punishments

        Exit Function
ErrorHandler:
106     Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SendUserPunishmentsDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT user_id, number, reason FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", UCase$(UserName))
    
102     If RS Is Nothing Then Exit Sub

104     If Not RS.RecordCount = 0 Then

108         While Not RS.EOF
110             Call WriteConsoleMsg(UserIndex, RS!Number & " - " & RS!Reason, e_FontTypeNames.FONTTYPE_INFO)
            
112             RS.MoveNext
            Wend

        End If

        Exit Sub
ErrorHandler:
114     Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetUserGuildIndexDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildIndexDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildMemberDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildMemberDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_member_history"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildAspirantDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildAspirantDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_aspirant_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildPedidosDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildPedidosDatabase = SanitizeNullValue(GetUserValue(LCase$(UserName), "guild_requests_history"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(UserName As String, Reason As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_rejected_because", Reason)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_index", GuildIndex)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_aspirant_index", AspirantIndex)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_member_history", guilds)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_requests_history", Pedidos)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim gName       As String

        Dim Miembro     As String

        Dim GuildActual As Integer

        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE UPPER(name) = ?;", UCase$(UserName))

102     If RS Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "Pj Inexistente", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Get the character's current guild
106     GuildActual = SanitizeNullValue(RS!Guild_Index, 0)

108     If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
110         gName = "<" & GuildName(GuildActual) & ">"
        Else
112         gName = "Ninguno"

        End If

        'Get previous guilds
114     Miembro = SanitizeNullValue(RS!guild_member_history, vbNullString)

116     If Len(Miembro) > 400 Then
118         Miembro = ".." & Right$(Miembro, 400)

        End If

120     Call WriteCharacterInfo(UserIndex, UserName, RS!race_id, RS!class_id, RS!genre_id, RS!level, RS!gold, RS!bank_gold, SanitizeNullValue(RS!guild_requests_history, vbNullString), gName, Miembro, RS!pertenece_real, RS!pertenece_caos, RS!ciudadanos_matados, RS!criminales_matados)

        Exit Sub
ErrorHandler:
122     Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function EnterAccountDatabase(ByVal UserIndex As Integer, ByVal CuentaEmail As String) As Boolean

        On Error GoTo ErrorHandler
    
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT id from account WHERE email = ?", UCase$(CuentaEmail))
    
102     If Connection.State = adStateClosed Then
104         Call WriteShowMessageBox(UserIndex, "Ha ocurrido un error interno en el servidor. ¡Estamos tratando de resolverlo!")
            Exit Function
        End If
    
122     UserList(UserIndex).AccountID = RS!ID
124     UserList(UserIndex).Cuenta = CuentaEmail
        UserList(UserIndex).Email = CuentaEmail
    
128     EnterAccountDatabase = True
    
        Exit Function

ErrorHandler:
130     Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function PersonajePerteneceID(ByVal UserName As String, ByVal AccountID As Long) As Boolean
    
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT id FROM user WHERE name = ? AND account_id = ?;", UserName, AccountID)
    
102     If RS Is Nothing Then
104         PersonajePerteneceID = False
            Exit Function
        End If
    
106     PersonajePerteneceID = True
    
End Function

Public Function SetPositionDatabase(UserName As String, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        On Error GoTo ErrorHandler

102     SetPositionDatabase = Execute("UPDATE user SET pos_map = ?, pos_x = ?, pos_y = ? WHERE UPPER(name) = ?;", Map, X, Y, UCase$(UserName))

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetMapDatabase(UserName As String) As Integer
        On Error GoTo ErrorHandler

100     GetMapDatabase = val(GetUserValue(LCase$(UserName), "pos_map"))

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function AddOroBancoDatabase(UserName As String, ByVal OroGanado As Long) As Boolean
        On Error GoTo ErrorHandler

102     AddOroBancoDatabase = Execute("UPDATE user SET bank_gold = bank_gold + ? WHERE UPPER(name) = ?;", OroGanado, UCase$(UserName))

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveAUsuarioDatabase(UserName As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

102     DarLlaveAUsuarioDatabase = Execute("INSERT INTO house_key (key_obj, account_id) values (?, (SELECT account_id FROM user WHERE UPPER(name) = ?))", LlaveObj, UCase$(UserName))

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveAUsuarioDatabase. UserName: " & UserName & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveACuentaDatabase(Email As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler
        'Hacer verificacion de que si alguien tiene esta llave, si alguien la tiene hay que prevenir la creacion.
' 101     LlaveYaOtorgadaAJugador = Execute("SELECT * FROM house_key WHERE key_obj = ?;", LlaveObj, UCase$(Email))
        
102     DarLlaveACuentaDatabase = Execute("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT id FROM account WHERE email = ?);", LlaveObj, UCase$(Email))
        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveACuentaDatabase. Email: " & Email & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function SacarLlaveDatabase(ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

        Dim i As Integer
        Dim UserCount As Integer
        Dim Users() As String

        ' Obtengo los usuarios logueados en la cuenta del dueño de la llave
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT name FROM `user` INNER JOIN `account` ON `user`.account_id = account.id INNER JOIN `house_key` ON `house_key`.account_id = account.id WHERE `house_key`.key_obj = ?;", LlaveObj)
    
102     If RS Is Nothing Then Exit Function

        ' Los almaceno en un array
104     UserCount = RS.RecordCount
    
106     ReDim Users(1 To UserCount) As String

110     i = 1

112     While Not RS.EOF
    
114         Users(i) = RS!Name
116         i = i + 1

118         RS.MoveNext
        Wend
    
        ' Intento borrar la llave de la db
120     SacarLlaveDatabase = Execute("DELETE FROM house_key WHERE key_obj = ?;", LlaveObj)
    
        ' Si pudimos borrar, actualizamos los usuarios logueados
        If (SacarLlaveDatabase) Then
            Dim UserIndex As Integer
        
122         For i = 1 To UserCount
124             UserIndex = NameIndex(Users(i))
            
126             If UserIndex <> 0 Then
128                 Call SacarLlaveDeLLavero(UserIndex, LlaveObj)
                End If
            Next
        End If
        
        Exit Function

ErrorHandler:
132     Call LogDatabaseError("Error in SacarLlaveDatabase. LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub VerLlavesDatabase(ByVal UserIndex As Integer)
        On Error GoTo ErrorHandler

        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT email, key_obj FROM `house_key` INNER JOIN `account` ON `house_key`.account_id = `account`.id;")

102     If RS Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", e_FontTypeNames.FONTTYPE_INFO)

106     ElseIf RS.RecordCount = 0 Then
108         Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", e_FontTypeNames.FONTTYPE_INFO)
    
        Else
            Dim Message As String
        
110         Message = "Llaves usadas: " & RS.RecordCount & vbNewLine

114         While Not RS.EOF
116             Message = Message & "Llave: " & RS!key_obj & " - Cuenta: " & RS!Email & vbNewLine

118             RS.MoveNext
            Wend
        
120         Message = Left$(Message, Len(Message) - 2)
        
122         Call WriteConsoleMsg(UserIndex, Message, e_FontTypeNames.FONTTYPE_INFO)
        End If

        Exit Sub

ErrorHandler:
124     Call LogDatabaseError("Error in VerLlavesDatabase. UserName: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

        
        Exit Function

SanitizeNullValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SanitizeNullValue", Erl)

        
End Function

Public Sub SetMessageInfoDatabase(ByVal Name As String, ByVal Message As String)
    Call Execute("update user set message_info = concat(message_info, ?) where upper(name) = ?;", Message, UCase$(Name))
End Sub

Public Sub ChangeNameDatabase(ByVal CurName As String, ByVal NewName As String)
    Call SetUserValue(CurName, "name", NewName)
End Sub

' Función auxiliar equivalente a la expresión "i++"
Private Function PostInc(Value As Integer) As Integer
100     PostInc = Value
102     Value = Value + 1
End Function

Public Sub ResetLastLogout()
    Call Execute("Update user set last_logout = 0")
End Sub
