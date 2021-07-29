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
Public Database_DataSource  As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Private Database_Connection As ADODB.Connection
Private Command             As ADODB.Command
Public QueryData            As ADODB.Recordset
Private RecordsAffected     As Long

Private Database_Connection_Async As ADODB.Connection
Private Database_Async_Queue      As Collection

Private QueryBuilder        As cStringBuilder
Private ConnectedOnce       As Boolean

Public Sub Database_Connect()

        '************************************************************************************
        'Author: Juan Andres Dalmasso
        'Last Modification: 17/10/2020
        '21/09/2019 Jopi - Agregue soporte a conexion via DSN. Solo para usuarios avanzados.
        '17/10/2020 WyroX - Agrego soporte a multiples statements en la misma query
        '************************************************************************************
        On Error GoTo ErrorHandler
        
        Dim ConnectionID As String
        
100     Set Database_Connection = New ADODB.Connection
        Set Database_Connection_Async = frmMain.CreateDatabaseAsync()
        
102     If Len(Database_DataSource) <> 0 Then
    
104         ConnectionID = "DATA SOURCE=" & Database_DataSource & ";"

        Else
    
106         ConnectionID = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & _
                                                   "SERVER=" & Database_Host & ";" & _
                                                   "DATABASE=" & Database_Name & ";" & _
                                                   "USER=" & Database_Username & ";" & _
                                                   "PASSWORD=" & Database_Password & ";" & _
                                                   "OPTION=3;MULTI_STATEMENTS=1"
        End If
        
        Database_Connection.ConnectionString = ConnectionID
        Database_Connection_Async.ConnectionString = ConnectionID
        
        Set Database_Async_Queue = New Collection
        
108     Debug.Print Database_Connection.ConnectionString
    
110     Database_Connection.CursorLocation = adUseClient
        Database_Connection_Async.CursorLocation = adUseClient
        
112     Call Database_Connection.Open
        Call Database_Connection_Async.Open

113     Set QueryBuilder = New cStringBuilder

114     ConnectedOnce = True

        Exit Sub
    
ErrorHandler:
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description)
    
118     If Not ConnectedOnce Then
120         Call MsgBox("No se pudo conectar a la base de datos. Mas información en logs/Database.log", vbCritical, "OBDC - Error")
122         End
        End If

End Sub

Public Sub Database_Close()

        '***************************************************
        'Author: Juan Andres Dalmasso
        'Last Modification: 18/09/2018
        'Nota by WyroX: Cerrar la conexion tambien libera
        'los recursos y cierra los RecordSet generados
        '***************************************************
        On Error GoTo ErrorHandler
    
100     Set Command = Nothing
        
102     If Database_Connection.State <> adStateClosed Then
104         Call Database_Connection.Close
        End If
        
        If Database_Connection_Async.State <> adStateClosed Then
            Call Database_Connection_Async.Close
        End If
        
106     Set Database_Connection = Nothing
        Set Database_Connection_Async = Nothing
        
        Exit Sub
     
ErrorHandler:
108     Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.Description)

End Sub


Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

        On Error GoTo ErrorHandler
    
        Dim LoopC As Long
        Dim ParamC As Long
        Dim Params() As Variant
    
102     With UserList(UserIndex)
        
            Dim i As Integer
104         ReDim Params(45)

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
144         Params(PostInc(i)) = .Invent.NroItems
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
        
198         Call MakeQuery(QUERY_SAVE_MAINPJ, False, Params)

            ' Para recibir el ID del user
200         Call MakeQuery("SELECT LAST_INSERT_ID();", False)

202         If QueryData Is Nothing Then
204             .ID = 1
            Else
206             .ID = val(QueryData.Fields(0).Value)
            End If
        
            ' ******************* ATRIBUTOS *******************
208         ReDim Params(NUMATRIBUTOS * 3 - 1)
210         ParamC = 0
        
212         For LoopC = 1 To NUMATRIBUTOS
214             Params(ParamC) = .ID
216             Params(ParamC + 1) = LoopC
218             Params(ParamC + 2) = .Stats.UserAtributos(LoopC)
            
220             ParamC = ParamC + 3
222         Next LoopC
        
224         Call MakeQuery(QUERY_SAVE_ATTRIBUTES, False, Params)
        
            ' ******************* SPELLS **********************
226         ReDim Params(MAXUSERHECHIZOS * 3 - 1)
228         ParamC = 0
        
230         For LoopC = 1 To MAXUSERHECHIZOS
232             Params(ParamC) = .ID
234             Params(ParamC + 1) = LoopC
236             Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
238             ParamC = ParamC + 3
240         Next LoopC

242         Call MakeQuery(QUERY_SAVE_SPELLS, False, Params)
        
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
        
264         Call MakeQuery(QUERY_SAVE_INVENTORY, False, Params)
        
            ' ******************* SKILLS *******************
266         ReDim Params(NUMSKILLS * 3 - 1)
268         ParamC = 0
        
270         For LoopC = 1 To NUMSKILLS
272             Params(ParamC) = .ID
274             Params(ParamC + 1) = LoopC
276             Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
278             ParamC = ParamC + 3
280         Next LoopC
        
282         Call MakeQuery(QUERY_SAVE_SKILLS, False, Params)
        
            ' ******************* QUESTS *******************
284         ReDim Params(MAXUSERQUESTS * 2 - 1)
286         ParamC = 0
        
288         For LoopC = 1 To MAXUSERQUESTS
290             Params(ParamC) = .ID
292             Params(ParamC + 1) = LoopC
            
294             ParamC = ParamC + 2
296         Next LoopC
        
298         Call MakeQuery(QUERY_SAVE_QUESTS, False, Params)
        
            ' ******************* PETS ********************
300         ReDim Params(MAXMASCOTAS * 3 - 1)
302         ParamC = 0
        
304         For LoopC = 1 To MAXMASCOTAS
306             Params(ParamC) = .ID
308             Params(ParamC + 1) = LoopC
310             Params(ParamC + 2) = 0
            
312             ParamC = ParamC + 3
314         Next LoopC
    
316         Call MakeQuery(QUERY_SAVE_PETS, False, Params)
    
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
    
100     Call QueryBuilder.Clear

        'Basic user data
102     With UserList(UserIndex)
        
104         ReDim Params(91)

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
150         Params(PostInc(i)) = .Invent.NroItems
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
204         Params(PostInc(i)) = .Stats.InventLevel
206         Params(PostInc(i)) = .Stats.ELO
208         Params(PostInc(i)) = .flags.Desnudo
210         Params(PostInc(i)) = .flags.Envenenado
212         Params(PostInc(i)) = .flags.Escondido
214         Params(PostInc(i)) = .flags.Hambre
216         Params(PostInc(i)) = .flags.Sed
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
238         Params(PostInc(i)) = (.flags.Privilegios And PlayerType.RoyalCouncil)
240         Params(PostInc(i)) = (.flags.Privilegios And PlayerType.ChaosCouncil)
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
278         Params(PostInc(i)) = IIf(Logout, 0, 1)
280         Params(PostInc(i)) = .Stats.Advertencias
282         Params(PostInc(i)) = .flags.ReturnPos.Map
284         Params(PostInc(i)) = .flags.ReturnPos.X
286         Params(PostInc(i)) = .flags.ReturnPos.Y

            ' WHERE block
288         Params(PostInc(i)) = .ID
            
            Call MakeQuery(QUERY_UPDATE_MAINPJ, True, Params)

            ' ************************** User attributes ****************************
302         If .flags.ModificoAttributos Then
304             ReDim Params(NUMATRIBUTOS * 3 - 1)
306             ParamC = 0
            
308             For LoopC = 1 To NUMATRIBUTOS
310                 Params(ParamC) = .ID
312                 Params(ParamC + 1) = LoopC
314                 Params(ParamC + 2) = .Stats.UserAtributosBackUP(LoopC)
                
316                 ParamC = ParamC + 3
318             Next LoopC
                
                Call MakeQuery(QUERY_UPSERT_ATTRIBUTES, True, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modificaron los atributos. WTF? Bueno, guardando..."
                .flags.ModificoAttributos = False
            End If

            ' ************************** User spells *********************************
332         If .flags.ModificoHechizos Then
334             ReDim Params(MAXUSERHECHIZOS * 3 - 1)
336             ParamC = 0
            
338             For LoopC = 1 To MAXUSERHECHIZOS
340                 Params(ParamC) = .ID
342                 Params(ParamC + 1) = LoopC
344                 Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
                
346                 ParamC = ParamC + 3
348             Next LoopC
                
                Call MakeQuery(QUERY_UPSERT_SPELLS, True, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modificaron los hechizos. Guardando..."
                .flags.ModificoHechizos = False
            End If
            
            ' ************************** User inventory *********************************
364         If .flags.ModificoInventario Then
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

                Call MakeQuery(QUERY_UPSERT_INVENTORY, True, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico el inventario. Guardando..."
                .flags.ModificoInventario = False
            End If
            
            ' ************************** User bank inventory *********************************
400         If .flags.ModificoInventarioBanco Then
402             ReDim Params(MAX_BANCOINVENTORY_SLOTS * 4 - 1)
404             ParamC = 0
            
406             For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
408                 Params(ParamC) = .ID
410                 Params(ParamC + 1) = LoopC
412                 Params(ParamC + 2) = .BancoInvent.Object(LoopC).ObjIndex
414                 Params(ParamC + 3) = .BancoInvent.Object(LoopC).amount
                
416                 ParamC = ParamC + 4
418             Next LoopC
    
                Call MakeQuery(QUERY_SAVE_BANCOINV, True, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico el inventario del banco. Guardando..."
                .flags.ModificoInventarioBanco = False
            End If

            ' ************************** User skills *********************************
434         If .flags.ModificoSkills Then
436             ReDim Params(NUMSKILLS * 3 - 1)
438             ParamC = 0
            
440             For LoopC = 1 To NUMSKILLS
442                 Params(ParamC) = .ID
444                 Params(ParamC + 1) = LoopC
446                 Params(ParamC + 2) = .Stats.UserSkills(LoopC)
                
448                 ParamC = ParamC + 3
450             Next LoopC
        
                Call MakeQuery(QUERY_UPSERT_SKILLS, True, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico las habilidades. Guardando..."
                .flags.ModificoSkills = False
            End If

            ' ************************** User pets *********************************
466         If .flags.ModificoMascotas Then
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
                
                Call MakeQuery(QUERY_UPSERT_PETS, True, Params)

                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico las mascotas. Guardando..."
                .flags.ModificoMascotas = False
            End If
        
            ' ************************** User connection logs *********************************
        
            'Agrego ip del user
            Call MakeQuery(QUERY_SAVE_CONNECTION, True, .ID, .IP)

            'Borro la mas vieja si hay mas de 5 (WyroX: si alguien sabe una forma mejor de hacerlo me avisa)
            Call MakeQuery(QUERY_DELETE_LAST_CONNECTIONS, True, .ID, .ID)

            ' ************************** User quests *********************************
524         If .flags.ModificoQuests Then
526             QueryBuilder.Append "INSERT INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
            
                Dim Tmp As Integer, LoopK As Long
    
528             For LoopC = 1 To MAXUSERQUESTS
530                 QueryBuilder.Append "("
532                 QueryBuilder.Append .ID & ", "
534                 QueryBuilder.Append LoopC & ", "
536                 QueryBuilder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
                
538                 If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
540                     Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs
    
542                     If Tmp Then
    
544                         For LoopK = 1 To Tmp
546                             QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                            
548                             If LoopK < Tmp Then
550                                 QueryBuilder.Append "-"
                                End If
    
552                         Next LoopK
                        
    
                        End If
    
                    End If
                
554                 QueryBuilder.Append "', '"
                
556                 If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                    
558                     Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                        
560                     For LoopK = 1 To Tmp
    
562                         QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                        
564                         If LoopK < Tmp Then
566                             QueryBuilder.Append "-"
                            End If
                    
568                     Next LoopK
                
                    End If
                
570                 QueryBuilder.Append "')"
    
572                 If LoopC < MAXUSERQUESTS Then
574                     QueryBuilder.Append ", "
                    End If
    
576             Next LoopC
            
578             QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id), npcs=VALUES(npcs);"
    
                Call MakeQuery(QueryBuilder.ToString(), True)

584             Call QueryBuilder.Clear
                
                ' Reseteamos el flag para no volver a guardar.
                Debug.Print "Se modifico las quests. Guardando..."
                .flags.ModificoQuests = False
            End If
        
            ' ************************** User completed quests *********************************
586         If .QuestStats.NumQuestsDone > 0 Then
                
588             If .flags.ModificoQuestsHechas Then
                
                    ' Armamos la query con los placeholders
590                 QueryBuilder.Append "INSERT INTO quest_done (user_id, quest_id) VALUES "
                
592                 For LoopC = 1 To .QuestStats.NumQuestsDone
594                     QueryBuilder.Append "(?, ?)"
                
596                     If LoopC < .QuestStats.NumQuestsDone Then
598                         QueryBuilder.Append ", "
                        End If
                
600                 Next LoopC
                        
602                 QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id); "
                
                    ' Metemos los parametros
604                 ReDim Params(.QuestStats.NumQuestsDone * 2 - 1)
606                 ParamC = 0
                
608                 For LoopC = 1 To .QuestStats.NumQuestsDone
610                     Params(ParamC) = .ID
612                     Params(ParamC + 1) = .QuestStats.QuestsDone(LoopC)
                    
614                     ParamC = ParamC + 2
616                 Next LoopC
        
622                 Call MakeQuery(QueryBuilder.ToString(), True, Params)

626                 Call QueryBuilder.Clear
                    
                    ' Reseteamos el flag para no volver a guardar.
                    Debug.Print "Se modifico las quests hechas. Guardando..."
                    .flags.ModificoQuestsHechas = False
                End If
                
            End If

            ' ************************** User logout *********************************
            ' Si deslogueó, actualizo la cuenta
628         If Logout Then
                Call MakeQuery("UPDATE account SET logged = logged - 1 WHERE id = ?; ", True, .AccountID)
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

102         Call MakeQuery(QUERY_LOAD_MAINPJ, False, .Name)

104         If QueryData Is Nothing Then Exit Sub

            'Start setting data
106         .ID = QueryData!ID
108         .Name = QueryData!Name
110         .Stats.ELV = QueryData!level
112         .Stats.Exp = QueryData!Exp
114         .genero = QueryData!genre_id
116         .raza = QueryData!race_id
118         .clase = QueryData!class_id
120         .Hogar = QueryData!home_id
122         .Desc = QueryData!Description
124         .Stats.GLD = QueryData!gold
126         .Stats.Banco = QueryData!bank_gold
128         .Stats.SkillPts = QueryData!free_skillpoints
130         .Pos.Map = QueryData!pos_map
132         .Pos.X = QueryData!pos_x
134         .Pos.Y = QueryData!pos_y
136         .MENSAJEINFORMACION = QueryData!message_info
138         .OrigChar.Body = QueryData!body_id
140         .OrigChar.Head = QueryData!head_id
142         .OrigChar.WeaponAnim = QueryData!weapon_id
144         .OrigChar.CascoAnim = QueryData!helmet_id
146         .OrigChar.ShieldAnim = QueryData!shield_id
148         .OrigChar.Heading = QueryData!Heading
150         .Invent.NroItems = QueryData!items_Amount
152         .Invent.ArmourEqpSlot = SanitizeNullValue(QueryData!slot_armour, 0)
154         .Invent.WeaponEqpSlot = SanitizeNullValue(QueryData!slot_weapon, 0)
156         .Invent.CascoEqpSlot = SanitizeNullValue(QueryData!slot_helmet, 0)
158         .Invent.EscudoEqpSlot = SanitizeNullValue(QueryData!slot_shield, 0)
160         .Invent.MunicionEqpSlot = SanitizeNullValue(QueryData!slot_ammo, 0)
162         .Invent.BarcoSlot = SanitizeNullValue(QueryData!slot_ship, 0)
164         .Invent.MonturaSlot = SanitizeNullValue(QueryData!slot_mount, 0)
166         .Invent.DañoMagicoEqpSlot = SanitizeNullValue(QueryData!slot_dm, 0)
168         .Invent.ResistenciaEqpSlot = SanitizeNullValue(QueryData!slot_rm, 0)
170         .Invent.NudilloSlot = SanitizeNullValue(QueryData!slot_knuckles, 0)
172         .Invent.HerramientaEqpSlot = SanitizeNullValue(QueryData!slot_tool, 0)
174         .Invent.MagicoSlot = SanitizeNullValue(QueryData!slot_magic, 0)
176         .Stats.MinHp = QueryData!min_hp
178         .Stats.MaxHp = QueryData!max_hp
180         .Stats.MinMAN = QueryData!min_man
182         .Stats.MaxMAN = QueryData!max_man
184         .Stats.MinSta = QueryData!min_sta
186         .Stats.MaxSta = QueryData!max_sta
188         .Stats.MinHam = QueryData!min_ham
190         .Stats.MaxHam = QueryData!max_ham
192         .Stats.MinAGU = QueryData!min_sed
194         .Stats.MaxAGU = QueryData!max_sed
196         .Stats.MinHIT = QueryData!min_hit
198         .Stats.MaxHit = QueryData!max_hit
200         .Stats.NPCsMuertos = QueryData!killed_npcs
202         .Stats.UsuariosMatados = QueryData!killed_users
204         .Stats.InventLevel = QueryData!invent_level
206         .Stats.ELO = QueryData!ELO
208         .flags.Desnudo = QueryData!is_naked
210         .flags.Envenenado = QueryData!is_poisoned
212         .flags.Escondido = QueryData!is_hidden
214         .flags.Hambre = QueryData!is_hungry
216         .flags.Sed = QueryData!is_thirsty
218         .flags.Ban = QueryData!is_banned
220         .flags.Muerto = QueryData!is_dead
222         .flags.Navegando = QueryData!is_sailing
224         .flags.Paralizado = QueryData!is_paralyzed
226         .flags.VecesQueMoriste = QueryData!deaths
228         .flags.Montado = QueryData!is_mounted
230         .flags.Pareja = QueryData!spouse
232         .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
234         .flags.Silenciado = QueryData!is_silenced
236         .flags.MinutosRestantes = QueryData!silence_minutes_left
238         .flags.SegundosPasados = QueryData!silence_elapsed_seconds
240         .flags.MascotasGuardadas = QueryData!pets_saved

246         .flags.ReturnPos.Map = QueryData!return_map
248         .flags.ReturnPos.X = QueryData!return_x
250         .flags.ReturnPos.Y = QueryData!return_y
        
252         .Counters.Pena = QueryData!counter_pena
        
254         .ChatGlobal = QueryData!chat_global
256         .ChatCombate = QueryData!chat_combate

258         If QueryData!pertenece_consejo_real Then
260             .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

            End If

262         If QueryData!pertenece_consejo_caos Then
264             .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

            End If

266         .Faccion.ArmadaReal = QueryData!pertenece_real
268         .Faccion.FuerzasCaos = QueryData!pertenece_caos
270         .Faccion.ciudadanosMatados = QueryData!ciudadanos_matados
272         .Faccion.CriminalesMatados = QueryData!criminales_matados
274         .Faccion.RecibioArmaduraReal = QueryData!recibio_armadura_real
276         .Faccion.RecibioArmaduraCaos = QueryData!recibio_armadura_caos
278         .Faccion.RecibioExpInicialReal = QueryData!recibio_exp_real
280         .Faccion.RecibioExpInicialCaos = QueryData!recibio_exp_caos
282         .Faccion.RecompensasReal = QueryData!recompensas_real
284         .Faccion.RecompensasCaos = QueryData!recompensas_caos
286         .Faccion.Reenlistadas = QueryData!Reenlistadas
288         .Faccion.NivelIngreso = SanitizeNullValue(QueryData!nivel_ingreso, 0)
290         .Faccion.MatadosIngreso = SanitizeNullValue(QueryData!matados_ingreso, 0)
292         .Faccion.NextRecompensa = SanitizeNullValue(QueryData!siguiente_recompensa, 0)
294         .Faccion.Status = QueryData!Status

296         .GuildIndex = SanitizeNullValue(QueryData!Guild_Index, 0)
        
298         .Stats.Advertencias = QueryData!warnings
        
            'User attributes
300         Call MakeQuery("SELECT value, number FROM attribute WHERE user_id = ?;", False, .ID)
    
302         If Not QueryData Is Nothing Then
304             QueryData.MoveFirst

306             While Not QueryData.EOF

308                 .Stats.UserAtributos(QueryData!Number) = QueryData!Value
310                 .Stats.UserAtributosBackUP(QueryData!Number) = .Stats.UserAtributos(QueryData!Number)

312                 QueryData.MoveNext
                Wend

            End If

            'User spells
314         Call MakeQuery("SELECT number, spell_id FROM spell WHERE user_id = ?;", False, .ID)

316         If Not QueryData Is Nothing Then
318             QueryData.MoveFirst

320             While Not QueryData.EOF

322                 .Stats.UserHechizos(QueryData!Number) = QueryData!spell_id

324                 QueryData.MoveNext
                Wend

            End If

            'User pets
326         Call MakeQuery("SELECT number, pet_id FROM pet WHERE user_id = ?;", False, .ID)

328         If Not QueryData Is Nothing Then
330             QueryData.MoveFirst

332             While Not QueryData.EOF

334                 .MascotasType(QueryData!Number) = QueryData!pet_id
                
336                 If val(QueryData!pet_id) <> 0 Then
338                     .NroMascotas = .NroMascotas + 1

                    End If

340                 QueryData.MoveNext
                Wend

            End If

            'User inventory
342         Call MakeQuery("SELECT number, item_id, is_equipped, amount FROM inventory_item WHERE user_id = ?;", False, .ID)

            
344         If Not QueryData Is Nothing Then
346             QueryData.MoveFirst

348             While Not QueryData.EOF

350                 With .Invent.Object(QueryData!Number)
352                     .ObjIndex = QueryData!item_id
                
354                     If .ObjIndex <> 0 Then
356                         If LenB(ObjData(.ObjIndex).Name) Then
358                             .amount = QueryData!amount
360                             .Equipped = QueryData!is_equipped
                            Else
362                             .ObjIndex = 0

                            End If

                        End If

                    End With

364                 QueryData.MoveNext
                Wend

            End If

            'User bank inventory
366         Call MakeQuery("SELECT number, item_id, amount FROM bank_item WHERE user_id = ?;", False, .ID)
            
            counter = 0
            
368         If Not QueryData Is Nothing Then
370             QueryData.MoveFirst

372             While Not QueryData.EOF

374                 With .BancoInvent.Object(QueryData!Number)
376                     .ObjIndex = QueryData!item_id
                
378                     If .ObjIndex <> 0 Then
380                         If LenB(ObjData(.ObjIndex).Name) Then
                                counter = counter + 1
                                
382                             .amount = QueryData!amount
                            Else
384                             .ObjIndex = 0

                            End If

                        End If

                    End With

386                 QueryData.MoveNext
                Wend
                
                .BancoInvent.NroItems = counter
            End If
            
            'User skills
388         Call MakeQuery("SELECT number, value FROM skillpoint WHERE user_id = ?;", False, .ID)

390         If Not QueryData Is Nothing Then
392             QueryData.MoveFirst

394             While Not QueryData.EOF

396                 .Stats.UserSkills(QueryData!Number) = QueryData!Value
                    '.Stats.ExpSkills(QueryData!Number) = QueryData!Exp
                    '.Stats.EluSkills(QueryData!Number) = QueryData!ELU

398                 QueryData.MoveNext
                Wend

            End If

            'User friends
            'Call MakeQuery("SELECT * FROM friend WHERE user_id = ?;", False, .Id)

            'If Not QueryData Is Nothing Then
            '    QueryData.MoveFirst

            '    While Not QueryData.EOF

            '200     .Amigos(QueryData!Number).Nombre = QueryData!friend
            '200     .Amigos(QueryData!Number).Ignorado = QueryData!Ignored

            '        QueryData.MoveNext
            '    Wend
            'End If
        
            Dim LoopC As Byte
        
            'User quests
400         Call MakeQuery("SELECT number, quest_id, npcs, npcstarget FROM quest WHERE user_id = ?;", False, .ID)

402         If Not QueryData Is Nothing Then
404             QueryData.MoveFirst

406             While Not QueryData.EOF

408                 .QuestStats.Quests(QueryData!Number).QuestIndex = QueryData!quest_id
                
410                 If .QuestStats.Quests(QueryData!Number).QuestIndex > 0 Then
412                     If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs Then

                            Dim NPCs() As String

414                         NPCs = Split(QueryData!NPCs, "-")
416                         ReDim .QuestStats.Quests(QueryData!Number).NPCsKilled(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs)

418                         For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs
420                             .QuestStats.Quests(QueryData!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
422                         Next LoopC

                        End If
                    
424                     If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs Then

                            Dim NPCsTarget() As String

426                         NPCsTarget = Split(QueryData!NPCsTarget, "-")
428                         ReDim .QuestStats.Quests(QueryData!Number).NPCsTarget(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs)

430                         For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs
432                             .QuestStats.Quests(QueryData!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
434                         Next LoopC

                        End If

                    End If

436                 QueryData.MoveNext
                Wend

            End If
        
            'User quests done
438         Call MakeQuery("SELECT quest_id FROM quest_done WHERE user_id = ?;", False, .ID)

440         If Not QueryData Is Nothing Then
442             .QuestStats.NumQuestsDone = QueryData.RecordCount
                
444             ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
        
446             QueryData.MoveFirst
            
448             LoopC = 1

450             While Not QueryData.EOF
            
452                 .QuestStats.QuestsDone(LoopC) = QueryData!quest_id
454                 LoopC = LoopC + 1

456                 QueryData.MoveNext
                Wend

            End If
        
            'User mail
            'TODO:
        
            ' Llaves
458         Call MakeQuery("SELECT key_obj FROM house_key WHERE account_id = ?", False, .AccountID)

460         If Not QueryData Is Nothing Then
462             QueryData.MoveFirst

464             LoopC = 1

466             While Not QueryData.EOF

468                 .Keys(LoopC) = QueryData!key_obj
470                 LoopC = LoopC + 1

472                 QueryData.MoveNext
                Wend

            End If

        End With

        Exit Sub

ErrorHandler:
474     Call LogDatabaseError("Error en LoadUserDatabase: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)



End Sub

Public Function MakeQuery(query As String, ByVal NoResultAndAsync As Boolean, ParamArray Query_Parameters() As Variant) As Boolean
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Hace una unica query a la db. Asume una conexion.
        ' Si NoResult = False, el metodo lee el resultado de la query
        ' Guarda el resultado en QueryData
    
        On Error GoTo ErrorHandler
        
        Dim ArgumentOuter As Variant
        Dim ArgumentInner As Variant
        
        If frmMain.chkLogDbPerfomance.Value = 1 Then Call GetElapsedTime

100     Set Command = New ADODB.Command
    
102     With Command

            If (NoResultAndAsync) Then
                .ActiveConnection = Database_Connection_Async
            Else
                .ActiveConnection = Database_Connection
            End If
            
110         .CommandText = query
106         .CommandType = adCmdText
            .Prepared = True
            
            For Each ArgumentOuter In Query_Parameters
                If (IsArray(ArgumentOuter)) Then
                    For Each ArgumentInner In ArgumentOuter
                        .Parameters.Append CreateParameter(ArgumentInner, adParamInput)
                    Next ArgumentInner
                Else
                    .Parameters.Append CreateParameter(ArgumentOuter, adParamInput)
                End If
            Next ArgumentOuter

124         If NoResultAndAsync Then
                Call Database_Async_Queue.Add(Command)

                If (Database_Async_Queue.Count = 1) Then
                    Call .Execute(, , adExecuteNoRecords + adAsyncExecute)
                End If
            Else
                Set QueryData = .Execute(RecordsAffected)
                
                If QueryData.State <> 0 Then
                    If QueryData.BOF Or QueryData.EOF Then
                        Set QueryData = Nothing
                    End If
                End If
            End If
        
        End With
        If frmMain.chkLogDbPerfomance.Value = 1 Then Call LogPerformance("Query: " & query & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    
        Exit Function
    
ErrorHandler:

        Dim errNumber As Long, ErrDesc As String
134     errNumber = Err.Number
136     ErrDesc = Err.Description

138     If Not adoIsConnected(Database_Connection) Then
140         Call LogDatabaseError("Alerta en MakeQuery: Se perdió la conexión con la DB. Reconectando.")
142         Call Database_Connect
144         Resume
        
        Else
146         Call LogDatabaseError("Error en MakeQuery: query = '" & query & "'. " & errNumber & " - " & ErrDesc)
        
            On Error GoTo 0

148         Err.raise errNumber, "MakeQuery", ErrDesc

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

Public Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para leer un unico valor de una unica fila

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = ?;", False, ValueTest)
    
        'Revisamos si recibio un resultado
102     If QueryData Is Nothing Then Exit Function

        'Obtenemos la variable
104     GetDBValue = QueryData.Fields(0).Value

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetCuentaValue(CuentaEmail As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor de la cuenta
        
        On Error GoTo GetCuentaValue_Err
        
100     GetCuentaValue = GetDBValue("account", Columna, "email", LCase$(CuentaEmail))

        
        Exit Function

GetCuentaValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetCuentaValue", Erl)

        
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor del char
        
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)

        
        Exit Function

GetUserValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserValue", Erl)

        
End Function

Public Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para escribir un unico valor de una unica fila

        On Error GoTo ErrorHandler

        Call MakeQuery("UPDATE " & Tabla & " SET " & ColumnaSet & " = ? WHERE " & ColumnaTest & " = ?;", True, ValueSet, ValueTest)
        
        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.Description)

End Sub

Private Sub SetCuentaValue(CuentaEmail As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        
        On Error GoTo SetCuentaValue_Err
        
100     Call SetDBValue("account", Columna, Value, "email", LCase$(CuentaEmail))

        
        Exit Sub

SetCuentaValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetCuentaValue", Erl)

        
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, Value, "name", CharName)

        
        Exit Sub

SetUserValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValue", Erl)

        
End Sub

Private Sub SetCuentaValueByID(ByVal AccountID As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        ' Por ID
        
        On Error GoTo SetCuentaValueByID_Err
        
100     Call SetDBValue("account", Columna, Value, "id", AccountID)

        
        Exit Sub

SetCuentaValueByID_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetCuentaValueByID", Erl)

        
End Sub

Private Sub SetUserValueByID(ByVal ID As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        ' Por ID
        
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, Value, "id", ID)

        
        Exit Sub

SetUserValueByID_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValueByID", Erl)

        
End Sub

Public Function CheckUserExists(Name As String) As Boolean
        
        On Error GoTo CheckUserExists_Err
        
100     CheckUserExists = GetUserValue(Name, "COUNT(*)") > 0

        
        Exit Function

CheckUserExists_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.CheckUserExists", Erl)

        
End Function

Public Function CheckCuentaExiste(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckCuentaExiste_Err
        
100     CheckCuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0

        
        Exit Function

CheckCuentaExiste_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.CheckCuentaExiste", Erl)

        
End Function

Public Function BANCheckDatabase(Name As String) As Boolean
        
        On Error GoTo BANCheckDatabase_Err
        
100     BANCheckDatabase = CBool(GetUserValue(Name, "is_banned"))

        
        Exit Function

BANCheckDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.BANCheckDatabase", Erl)

        
End Function

Public Function GetCodigoActivacionDatabase(Name As String) As String
        
        On Error GoTo GetCodigoActivacionDatabase_Err
        
100     GetCodigoActivacionDatabase = GetCuentaValue(Name, "validate_code")

        
        Exit Function

GetCodigoActivacionDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetCodigoActivacionDatabase", Erl)

        
End Function

Public Function CheckCuentaActivadaDatabase(Name As String) As Boolean
        
        On Error GoTo CheckCuentaActivadaDatabase_Err
        
100     CheckCuentaActivadaDatabase = GetCuentaValue(Name, "validated")

        
        Exit Function

CheckCuentaActivadaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.CheckCuentaActivadaDatabase", Erl)

        
End Function

Public Function GetEmailDatabase(Name As String) As String
        
        On Error GoTo GetEmailDatabase_Err
        
100     GetEmailDatabase = GetCuentaValue(Name, "email")

        
        Exit Function

GetEmailDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetEmailDatabase", Erl)

        
End Function

Public Function GetMacAddressDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMacAddressDatabase_Err
        
100     GetMacAddressDatabase = GetCuentaValue(CuentaEmail, "mac_address")

        
        Exit Function

GetMacAddressDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetMacAddressDatabase", Erl)

        
End Function

Public Function GetHDSerialDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetHDSerialDatabase_Err
        
100     GetHDSerialDatabase = GetCuentaValue(CuentaEmail, "hd_serial")

        
        Exit Function

GetHDSerialDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetHDSerialDatabase", Erl)

        
End Function

Public Function CheckBanCuentaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckBanCuentaDatabase_Err
        
100     CheckBanCuentaDatabase = CBool(GetCuentaValue(CuentaEmail, "is_banned"))

        
        Exit Function

CheckBanCuentaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.CheckBanCuentaDatabase", Erl)

        
End Function

Public Function GetMotivoBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMotivoBanCuentaDatabase_Err
        
100     GetMotivoBanCuentaDatabase = GetCuentaValue(CuentaEmail, "ban_reason")

        
        Exit Function

GetMotivoBanCuentaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetMotivoBanCuentaDatabase", Erl)

        
End Function

Public Function GetQuienBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetQuienBanCuentaDatabase_Err
        
100     GetQuienBanCuentaDatabase = GetCuentaValue(CuentaEmail, "banned_by")

        
        Exit Function

GetQuienBanCuentaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetQuienBanCuentaDatabase", Erl)

        
End Function

Public Function GetCuentaLogeadaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo GetCuentaLogeadaDatabase_Err
        
100     GetCuentaLogeadaDatabase = GetCuentaValue(CuentaEmail, "is_logged")

        
        Exit Function

GetCuentaLogeadaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetCuentaLogeadaDatabase", Erl)

        
End Function

Public Function GetUserStatusDatabase(Name As String) As Integer
        
        On Error GoTo GetUserStatusDatabase_Err
        
100     GetUserStatusDatabase = GetUserValue(Name, "status")

        
        Exit Function

GetUserStatusDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserStatusDatabase", Erl)

        
End Function

Public Function GetAccountIDDatabase(Name As String) As Long
        
        On Error GoTo GetAccountIDDatabase_Err
        

        Dim Temp As Variant

100     Temp = GetUserValue(Name, "account_id")
    
102     If VBA.IsEmpty(Temp) Then
104         GetAccountIDDatabase = -1
        Else
106         GetAccountIDDatabase = val(Temp)

        End If

        
        Exit Function

GetAccountIDDatabase_Err:
108     Call TraceError(Err.Number, Err.Description, "modDatabase.GetAccountIDDatabase", Erl)

        
End Function

Public Sub GetPasswordAndSaltDatabase(CuentaEmail As String, PasswordHash As String, Salt As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT password, salt FROM account WHERE deleted = FALSE AND email = ?;", False, LCase$(CuentaEmail))

102     If QueryData Is Nothing Then Exit Sub
    
104     PasswordHash = QueryData!Password
106     Salt = QueryData!Salt
    
        Exit Sub
    
ErrorHandler:
108     Call LogDatabaseError("Error in GetPasswordAndSaltDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Function GetPersonajesCountDatabase(CuentaEmail As String) As Byte

        On Error GoTo ErrorHandler

        Dim ID As Long

100     ID = GetDBValue("account", "id", "email", LCase$(CuentaEmail))
    
102     GetPersonajesCountDatabase = GetPersonajesCountByIDDatabase(ID)
    
        Exit Function
    
ErrorHandler:
104     Call LogDatabaseError("Error in GetPersonajesCountDatabase. name: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountID As Long) As Byte

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT COUNT(*) FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountID)
    
102     If QueryData Is Nothing Then Exit Function
    
104     GetPersonajesCountByIDDatabase = QueryData.Fields(0).Value
    
        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountID As Long, Personaje() As PersonajeCuenta) As Byte
        
        On Error GoTo GetPersonajesCuentaDatabase_Err
        

100     Call MakeQuery("SELECT name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountID)

102     If QueryData Is Nothing Then Exit Function
    
104     GetPersonajesCuentaDatabase = QueryData.RecordCount
        
106     QueryData.MoveFirst
    
        Dim i As Integer

108     For i = 1 To GetPersonajesCuentaDatabase
110         Personaje(i).nombre = QueryData!Name
112         Personaje(i).Cabeza = QueryData!head_id
114         Personaje(i).clase = QueryData!class_id
116         Personaje(i).cuerpo = QueryData!body_id
118         Personaje(i).Mapa = QueryData!pos_map
120         Personaje(i).posX = QueryData!pos_x
122         Personaje(i).posY = QueryData!pos_y
124         Personaje(i).nivel = QueryData!level
126         Personaje(i).Status = QueryData!Status
128         Personaje(i).Casco = QueryData!helmet_id
130         Personaje(i).Escudo = QueryData!shield_id
132         Personaje(i).Arma = QueryData!weapon_id
134         Personaje(i).ClanIndex = QueryData!Guild_Index
        
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

156         If val(QueryData!is_dead) = 1 Or val(QueryData!is_sailing) = 1 Then
158             Personaje(i).Cabeza = 0
            End If
        
160         QueryData.MoveNext
        Next

        
        Exit Function

GetPersonajesCuentaDatabase_Err:
162     Call TraceError(Err.Number, Err.Description, "modDatabase.GetPersonajesCuentaDatabase", Erl)

        
End Function

Public Sub SetUserLoggedDatabase(ByVal ID As Long, ByVal AccountID As Long)
        
        On Error GoTo SetUserLoggedDatabase_Err
        
100     Call SetDBValue("user", "is_logged", 1, "id", ID)
        
        Call MakeQuery("UPDATE account SET logged = logged + 1 WHERE id = ?", True, AccountID)

        Exit Sub

SetUserLoggedDatabase_Err:
104     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserLoggedDatabase", Erl)

        
End Sub

Public Sub ResetLoggedDatabase(ByVal AccountID As Long)
        
        On Error GoTo ResetLoggedDatabase_Err
        
        Call MakeQuery("UPDATE account SET logged = 0 WHERE id = ?", True, AccountID)
        
        Exit Sub

ResetLoggedDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.ResetLoggedDatabase", Erl)

        
End Sub

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
        
        On Error GoTo SetUsersLoggedDatabase_Err
        
        Call MakeQuery("UPDATE statistics SET value = ? WHERE name = 'online'", True, NumUsers)
        
        Exit Sub

SetUsersLoggedDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUsersLoggedDatabase", Erl)

        
End Sub

Public Function LeerRecordUsuariosDatabase() As Long
        
        On Error GoTo LeerRecordUsuariosDatabase_Err
        
100     Call MakeQuery("SELECT value FROM statistics WHERE name = 'record';", False)

102     If QueryData Is Nothing Then Exit Function

104     LeerRecordUsuariosDatabase = val(QueryData!Value)

        Exit Function

LeerRecordUsuariosDatabase_Err:
106     Call TraceError(Err.Number, Err.Description, "modDatabase.LeerRecordUsuariosDatabase", Erl)

        
End Function

Public Sub SetRecordUsersDatabase(ByVal Record As Long)
        
        On Error GoTo SetRecordUsersDatabase_Err
        
        Call MakeQuery("UPDATE statistics SET value = ? WHERE name = 'record'", True, CStr(Record))
        
        Exit Sub

SetRecordUsersDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetRecordUsersDatabase", Erl)

        
End Sub

Public Sub LogoutAllUsersAndAccounts()
        
        On Error GoTo LogoutAllUsersAndAccounts_Err

100     Call MakeQuery("UPDATE user SET is_logged = 0;", True)
102     Call MakeQuery("UPDATE account SET logged = 0;", True)
        
        Exit Sub

LogoutAllUsersAndAccounts_Err:
104     Call TraceError(Err.Number, Err.Description, "modDatabase.LogoutAllUsersAndAccounts", Erl)

        
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
        
        Call MakeQuery("UPDATE skillpoints SET value = ? WHERE number = ? AND user_id = (SELECT id FROM user WHERE UPPER(name) = ?)", True, Value, Skill, UCase$(UserName))
        
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

Public Sub SaveNewAccountDatabase(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("INSERT INTO account SET email = ?, password = ?, salt = ?, validate_code = ?, date_created = NOW();", True, LCase$(CuentaEmail), PasswordHash, Salt, Codigo)
    
        Exit Sub
        
ErrorHandler:
102     Call LogDatabaseError("Error en SaveNewAccountDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub ValidarCuentaDatabase(UserCuenta As String)
        
        On Error GoTo ValidarCuentaDatabase_Err
        
100     Call SetCuentaValue(UserCuenta, "validated", 1)
        
        Exit Sub

ValidarCuentaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.ValidarCuentaDatabase", Erl)

        
End Sub

Public Function CheckUserAccount(Name As String, ByVal AccountID As Long) As Boolean

100     CheckUserAccount = (val(GetUserValue(Name, "account_id")) = AccountID)

End Function

Public Sub BorrarUsuarioDatabase(Name As String)

        On Error GoTo ErrorHandler
        
        Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE UPPER(name) = ?", True, UCase$(Name))

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en BorrarUsuarioDatabase borrando user de la Mysql Database: " & Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub BorrarCuentaDatabase(CuentaEmail As String)

        On Error GoTo ErrorHandler

        Dim ID As Long

100     ID = GetDBValue("account", "id", "email", LCase$(CuentaEmail))
        
        Call MakeQuery("UPDATE account SET email = CONCAT('DELETED_', email), deleted = TRUE WHERE email = ?", True, UCase$(CuentaEmail))
        Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE account_id = ?", True, ID)
  
        Exit Sub
    
ErrorHandler:
106     Call LogDatabaseError("Error en BorrarCuentaDatabase borrando user de la Mysql Database: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanDatabase(UserName As String, Reason As String, BannedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Call MakeQuery("UPDATE user SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE UPPER(name) = ?", True, BannedBy, Reason, UCase$(UserName))
        
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
        
        Call MakeQuery("UPDATE user SET warnings = warnings + 1 WHERE UPPER(name) = ?", True, UCase$(UserName))
        
102     Call SavePenaDatabase(UserName, "Advertencia de: " & WarnedBy & " debido a " & Reason)
    
    Exit Sub

ErrorHandler:
104     Call LogDatabaseError("Error in SaveWarnDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SavePenaDatabase(UserName As String, Reason As String)

        On Error GoTo ErrorHandler

        Call QueryBuilder.Clear

        Dim query As String
100     query = "INSERT INTO punishment(user_id, NUMBER, reason)"
102     query = query & " SELECT u.id, COUNT(p.number) + 1, ? FROM user u LEFT JOIN punishment p ON p.user_id = u.id WHERE UPPER(u.name) = ? "
        
        Call MakeQuery(query, True, Reason, UCase$(UserName))

        Exit Sub

ErrorHandler:
106     Call LogDatabaseError("Error in SavePenaDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SilenciarUserDatabase(UserName As String, ByVal Tiempo As Integer)
    
        On Error GoTo ErrorHandler
        
        Call MakeQuery("UPDATE user SET is_silenced = 1, silence_minutes_left = ?, silence_elapsed_seconds = 0 WHERE UPPER(name) = ?", True, Tiempo, UCase$(UserName))
        
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
        
        Call MakeQuery("UPDATE user SET is_banned = FALSE, banned_by = '', ban_reason = '' WHERE UPPER(name) = ?", True, UCase$(UserName))
        
        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanCuentaDatabase(ByVal AccountID As Long, Reason As String, BannedBy As String)

        On Error GoTo ErrorHandler
        
        Call MakeQuery("UPDATE account SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE id = ?", True, BannedBy, Reason, AccountID)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SaveBanCuentaDatabase: AccountId=" & AccountID & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub EcharConsejoDatabase(UserName As String)
        
        On Error GoTo EcharConsejoDatabase_Err
        
        Call MakeQuery("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE UPPER(name) = ?", True, UCase$(UserName))
        
        Exit Sub

EcharConsejoDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharConsejoDatabase", Erl)

        
End Sub

Public Sub EcharLegionDatabase(UserName As String)
        
        On Error GoTo EcharLegionDatabase_Err
        
        Call MakeQuery("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?", True, UCase$(UserName))
        
        Exit Sub

EcharLegionDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharLegionDatabase", Erl)

        
End Sub

Public Sub EcharArmadaDatabase(UserName As String)
        
        On Error GoTo EcharArmadaDatabase_Err
        
        Call MakeQuery("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?", True, UCase$(UserName))

        Exit Sub

EcharArmadaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharArmadaDatabase", Erl)

        
End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
        
        On Error GoTo CambiarPenaDatabase_Err
        
        Call MakeQuery("UPDATE punishment SET reason = ? WHERE number = ? AND user_id = (SELECT id from user WHERE UPPER(name) = ?)", True, Pena, Numero, UCase$(UserName))
        
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

100     Call MakeQuery("SELECT COUNT(*) as punishments FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", False, UCase$(UserName))

102     If QueryData Is Nothing Then Exit Function

104     GetUserAmountOfPunishmentsDatabase = QueryData!punishments

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

100     Call MakeQuery("SELECT user_id, number, reason FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", False, UCase$(UserName))
    
102     If QueryData Is Nothing Then Exit Sub

104     If Not QueryData.RecordCount = 0 Then
106         QueryData.MoveFirst

108         While Not QueryData.EOF
110             Call WriteConsoleMsg(UserIndex, QueryData!Number & " - " & QueryData!Reason, FontTypeNames.FONTTYPE_INFO)
            
112             QueryData.MoveNext
            Wend

        End If

        Exit Sub
ErrorHandler:
114     Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub


Public Function GetNombreCuentaDatabase(Name As String) As String

        On Error GoTo ErrorHandler

        'Hacemos la query.
100     Call MakeQuery("SELECT email FROM `account` INNER JOIN `user` ON user.account_id = account.id WHERE UPPER(user.name) = ?;", False, UCase$(Name))
    
        'Verificamos que la query no devuelva un resultado vacio.
102     If QueryData Is Nothing Then Exit Function
    
        'Obtenemos el nombre de la cuenta
104     GetNombreCuentaDatabase = QueryData!Email

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetNombreCuentaDatabase leyendo user de la Mysql Database: " & Name & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildIndexDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildIndexDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_index"), 0)

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

100     GetUserGuildMemberDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_member_history"), vbNullString)

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

100     GetUserGuildAspirantDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_aspirant_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildRejectionReasonDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_rejected_because"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildPedidosDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildPedidosDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_requests_history"), vbNullString)

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

100     Call MakeQuery("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE UPPER(name) = ?;", False, UCase$(UserName))

102     If QueryData Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Get the character's current guild
106     GuildActual = SanitizeNullValue(QueryData!Guild_Index, 0)

108     If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
110         gName = "<" & GuildName(GuildActual) & ">"
        Else
112         gName = "Ninguno"

        End If

        'Get previous guilds
114     Miembro = SanitizeNullValue(QueryData!guild_member_history, vbNullString)

116     If Len(Miembro) > 400 Then
118         Miembro = ".." & Right$(Miembro, 400)

        End If

120     Call WriteCharacterInfo(UserIndex, UserName, QueryData!race_id, QueryData!class_id, QueryData!genre_id, QueryData!level, QueryData!gold, QueryData!bank_gold, SanitizeNullValue(QueryData!guild_requests_history, vbNullString), gName, Miembro, QueryData!pertenece_real, QueryData!pertenece_caos, QueryData!ciudadanos_matados, QueryData!criminales_matados)

        Exit Sub
ErrorHandler:
122     Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function EnterAccountDatabase(ByVal UserIndex As Integer, CuentaEmail As String, Password As String, MacAddress As String, ByVal HDSerial As Long, IP As String) As Boolean

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT id, password, salt, validated, is_banned, ban_reason, banned_by FROM account WHERE email = ?;", False, LCase$(CuentaEmail))
    
102     If Database_Connection.State = adStateClosed Then
104         Call WriteShowMessageBox(UserIndex, "Ha ocurrido un error interno en el servidor. ¡Estamos tratando de resolverlo!")
            Exit Function
        End If
    
106     If QueryData Is Nothing Then
108         Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")
            Exit Function
        End If
    
110     If val(QueryData!is_banned) > 0 Then
112         Call WriteShowMessageBox(UserIndex, "La cuenta se encuentra baneada debido a: " & QueryData!ban_reason & ". Esta decisión fue tomada por: " & QueryData!banned_by & ".")
            Exit Function
        End If
    
114     If Not PasswordValida(Password, QueryData!Password, QueryData!Salt) Then
116         Call WriteShowMessageBox(UserIndex, "Contraseña inválida.")
            Exit Function
        End If
    
118     If val(QueryData!validated) = 0 Then
120         Call WriteShowMessageBox(UserIndex, "¡La cuenta no ha sido validada aún!")
            Exit Function
        End If
    
122     UserList(UserIndex).AccountID = QueryData!ID
124     UserList(UserIndex).Cuenta = CuentaEmail
        
        Call MakeQuery("UPDATE account SET mac_address = ?, hd_serial = ?, last_ip = ?, last_access = NOW() WHERE id = ?", True, MacAddress, HDSerial, IP, CLng(QueryData!ID))
        
128     EnterAccountDatabase = True
    
        Exit Function

ErrorHandler:
130     Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function PersonajePerteneceEmail(ByVal UserName As String, ByVal AccountEmail As String) As Boolean
    
100     Call MakeQuery("SELECT id FROM user INNER JOIN account ON user.account_id = account.id WHERE user.name = ? AND account.email = ?;", False, UserName, AccountEmail)
    
102     If QueryData Is Nothing Then
104         PersonajePerteneceEmail = False
            Exit Function
        End If
    
106     PersonajePerteneceEmail = True
    
End Function

Public Function PersonajePerteneceID(ByVal UserName As String, ByVal AccountID As Long) As Boolean
    
100     Call MakeQuery("SELECT id FROM user WHERE name = ? AND account_id = ?;", False, UserName, AccountID)
    
102     If QueryData Is Nothing Then
104         PersonajePerteneceID = False
            Exit Function
        End If
    
106     PersonajePerteneceID = True
    
End Function

Public Sub ChangePasswordDatabase(ByVal UserIndex As Integer, OldPassword As String, NewPassword As String)

        On Error GoTo ErrorHandler

100     If LenB(NewPassword) = 0 Then
102         Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
104     Call MakeQuery("SELECT password, salt FROM account WHERE id = ?;", False, UserList(UserIndex).AccountID)
    
106     If QueryData Is Nothing Then
108         Call WriteConsoleMsg(UserIndex, "No se ha podido cambiar la contraseña por un error interno. Avise a un administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
110     If Not PasswordValida(OldPassword, QueryData!Password, QueryData!Salt) Then
112         Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        Dim Salt As String * 10
114         Salt = RandomString(10) ' Alfanumerico
    
        Dim oSHA256 As CSHA256
116     Set oSHA256 = New CSHA256

        Dim PasswordHash As String * 64
118         PasswordHash = oSHA256.SHA256(NewPassword & Salt)
    
120     Set oSHA256 = Nothing
        
        Call MakeQuery("UPDATE account SET password = ?, salt = ? WHERE id = ?", True, PasswordHash, Salt, UserList(UserIndex).AccountID)

124     Call WriteConsoleMsg(UserIndex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
    
        Exit Sub

ErrorHandler:
126     Call LogDatabaseError("Error in ChangePasswordDatabase. Username: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetUsersLoggedAccountDatabase(ByVal AccountID As Long) As Byte

        On Error GoTo ErrorHandler

100     Call GetDBValue("account", "logged", "id", AccountID)
    
102     If QueryData Is Nothing Then Exit Function
    
104     GetUsersLoggedAccountDatabase = val(QueryData!logged)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUsersLoggedAccountDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function SetPositionDatabase(UserName As String, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET pos_map = ?, pos_x = ?, pos_y = ? WHERE UPPER(name) = ?;", False, Map, X, Y, UCase$(UserName))
    
102     SetPositionDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetMapDatabase(UserName As String) As Integer
        On Error GoTo ErrorHandler

100     GetMapDatabase = val(GetUserValue(UserName, "pos_map"))

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function AddOroBancoDatabase(UserName As String, ByVal OroGanado As Long) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET bank_gold = bank_gold + ? WHERE UPPER(name) = ?;", False, OroGanado, UCase$(UserName))
    
102     AddOroBancoDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveAUsuarioDatabase(UserName As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT account_id FROM user WHERE UPPER(name) = ?);", False, LlaveObj, UCase$(UserName))
    
102     DarLlaveAUsuarioDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveAUsuarioDatabase. UserName: " & UserName & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveACuentaDatabase(Email As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT id FROM account WHERE UPPER(email) = ?);", False, LlaveObj, UCase$(Email))
    
102     DarLlaveACuentaDatabase = RecordsAffected > 0
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
100     Call MakeQuery("SELECT name FROM `user` INNER JOIN `account` ON `user`.account_id = account.id INNER JOIN `house_key` ON `house_key`.account_id = account.id WHERE `user`.is_logged = TRUE AND `house_key`.key_obj = ?;", False, LlaveObj)
    
102     If QueryData Is Nothing Then Exit Function

        ' Los almaceno en un array
104     UserCount = QueryData.RecordCount
    
106     ReDim Users(1 To UserCount) As String
    
108     QueryData.MoveFirst

110     i = 1

112     While Not QueryData.EOF
    
114         Users(i) = QueryData!Name
116         i = i + 1

118         QueryData.MoveNext
        Wend
    
        ' Intento borrar la llave de la db
120     Call MakeQuery("DELETE FROM house_key WHERE key_obj = ?;", False, LlaveObj)
    
        ' Si pudimos borrar, actualizamos los usuarios logueados
        Dim UserIndex As Integer
    
122     For i = 1 To UserCount
124         UserIndex = NameIndex(Users(i))
        
126         If UserIndex <> 0 Then
128             Call SacarLlaveDeLLavero(UserIndex, LlaveObj)
            End If
        Next
    
130     SacarLlaveDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
132     Call LogDatabaseError("Error in SacarLlaveDatabase. LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub VerLlavesDatabase(ByVal UserIndex As Integer)
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT email, key_obj FROM `house_key` INNER JOIN `account` ON `house_key`.account_id = `account`.id;", False)

102     If QueryData Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", FontTypeNames.FONTTYPE_INFO)

106     ElseIf QueryData.RecordCount = 0 Then
108         Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", FontTypeNames.FONTTYPE_INFO)
    
        Else
            Dim Message As String
        
110         Message = "Llaves usadas: " & QueryData.RecordCount & vbNewLine
    
112         QueryData.MoveFirst

114         While Not QueryData.EOF
116             Message = Message & "Llave: " & QueryData!key_obj & " - Cuenta: " & QueryData!Email & vbNewLine

118             QueryData.MoveNext
            Wend
        
120         Message = Left$(Message, Len(Message) - 2)
        
122         Call WriteConsoleMsg(UserIndex, Message, FontTypeNames.FONTTYPE_INFO)
        End If

        Exit Sub

ErrorHandler:
124     Call LogDatabaseError("Error in VerLlavesDatabase. UserName: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetAccountID(Email As String) As String
    On Error GoTo ErrorHandler

100     GetAccountID = val(GetCuentaValue(Email, "id"))

        Exit Function
ErrorHandler:
102     Call LogDatabaseError("Error in GetAccountID. Email: " & Email & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetBaneoAccountId(ByVal AccountID As Long) As Boolean
100     GetBaneoAccountId = CBool(GetDBValue("account", "is_banned", "id", AccountID))
End Function

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

        
        Exit Function

SanitizeNullValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SanitizeNullValue", Erl)

        
End Function

Public Function GetUserLevelDatabase(ByVal Name As String) As Byte
        On Error GoTo GetUserLevelDatabase_Err

100     GetUserLevelDatabase = val(GetUserValue(Name, "level"))
        
        Exit Function
        
GetUserLevelDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserLevelDatabase", Erl)

End Function

Public Sub SetMessageInfoDatabase(ByVal Name As String, ByVal Message As String)
100     Call MakeQuery("update user set message_info = concat(message_info, ?) where upper(name) = ?;", True, Message, UCase$(Name))
End Sub

Public Sub ChangeNameDatabase(ByVal CurName As String, ByVal NewName As String)
    Call SetUserValue(CurName, "name", NewName)
End Sub

Public Sub OnDatabaseAsyncConnect(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If (Not pError Is Nothing) Then
        Call RegistrarError(pError.Number, pError.Description, pError.Source)
        End
    End If

    'Reinicio los users online
    Call SetUsersLoggedDatabase(0)
    
    'Leo el record de usuarios
    RecordUsuarios = LeerRecordUsuariosDatabase()
    
    'Tarea pesada
    Call LogoutAllUsersAndAccounts
End Sub

Public Sub OnDatabaseAsyncComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    If (Not pError Is Nothing) Then
        Call RegistrarError(pError.Number, pError.Description, pCommand.CommandText)
    End If

    Call Database_Async_Queue.Remove(1)
    
    If (Database_Async_Queue.Count >= 1) Then
        Dim Command As ADODB.Command
        Set Command = Database_Async_Queue.Item(1)
        
        Call Command.Execute(, , adExecuteNoRecords + adAsyncExecute)
    End If
End Sub

Function adoIsConnected(adoCn As ADODB.Connection) As Boolean

        '----------------------------------------------------------------
        '#PURPOSE: Checks whether the supplied db connection is alive and
        '          hasn't had it's TCP connection forcibly closed by remote
        '          host, for example, as happens during an undock event
        '#RETURNS: True if the supplied db is connected and error-free,
        '          False otherwise
        '#AUTHOR:  Belladonna
        '----------------------------------------------------------------

        ' No sacar
        On Error Resume Next

        Dim i As Long
        Dim cmd As New ADODB.Command

        'Set up SQL command to return 1
100     cmd.CommandText = "SELECT 1"
102     cmd.ActiveConnection = adoCn

        'Run a simple query, to test the connection
        
104     i = cmd.Execute.Fields(0)
        On Error GoTo 0

        'Tidy up
106     Set cmd = Nothing

        'If i is 1, connection is open
108     If i = 1 Then
110         adoIsConnected = True
        Else
112         adoIsConnected = False
        End If

End Function

' Función auxiliar equivalente a la expresión "i++"
Private Function PostInc(Value As Integer) As Integer
100     PostInc = Value
102     Value = Value + 1
End Function
