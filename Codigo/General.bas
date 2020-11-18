Attribute VB_Name = "General"

'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Public Type TDonador

    activo As Byte
    CreditoDonador As Integer
    FechaExpiracion As Date

End Type

Option Explicit

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Global LeerNPCs As New clsIniReader

Sub DarCuerpoDesnudo(ByVal Userindex As Integer)
        
        On Error GoTo DarCuerpoDesnudo_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/14/07
        'Da cuerpo desnudo a un usuario
        '***************************************************
        Dim CuerpoDesnudo As Integer

100     Select Case UserList(Userindex).genero

            Case eGenero.Hombre

102             Select Case UserList(Userindex).raza

                    Case eRaza.Humano
104                     CuerpoDesnudo = 21 'ok

106                 Case eRaza.Drow
108                     CuerpoDesnudo = 32 ' ok

110                 Case eRaza.Elfo
112                     CuerpoDesnudo = 510 'Revisar

114                 Case eRaza.Gnomo
116                     CuerpoDesnudo = 508 'Revisar

118                 Case eRaza.Enano
120                     CuerpoDesnudo = 53 'ok

122                 Case eRaza.Orco
124                     CuerpoDesnudo = 248 ' ok

                End Select

126         Case eGenero.Mujer

128             Select Case UserList(Userindex).raza

                    Case eRaza.Humano
130                     CuerpoDesnudo = 39 'ok

132                 Case eRaza.Drow
134                     CuerpoDesnudo = 40 'ok

136                 Case eRaza.Elfo
138                     CuerpoDesnudo = 511 'Revisar

140                 Case eRaza.Gnomo
142                     CuerpoDesnudo = 509 'Revisar

144                 Case eRaza.Enano
146                     CuerpoDesnudo = 60 ' ok

148                 Case eRaza.Orco
150                     CuerpoDesnudo = 249 'ok

                End Select

        End Select

152     UserList(Userindex).Char.Body = CuerpoDesnudo

154     UserList(Userindex).flags.Desnudo = 1

        
        Exit Sub

DarCuerpoDesnudo_Err:
        Call RegistrarError(Err.Number, Err.description, "General.DarCuerpoDesnudo", Erl)
        Resume Next
        
End Sub

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
        'b ahora es boolean,
        'b=true bloquea el tile en (x,y)
        'b=false desbloquea el tile en (x,y)
        'toMap = true -> Envia los datos a todo el mapa
        'toMap = false -> Envia los datos al user
        'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
        'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
        
        On Error GoTo Bloquear_Err
        

100     If toMap Then
102         Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
        Else
104         Call WriteBlockPosition(sndIndex, X, Y, b)

        End If

        
        Exit Sub

Bloquear_Err:
        Call RegistrarError(Err.Number, Err.description, "General.Bloquear", Erl)
        Resume Next
        
End Sub

Function HayCosta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayCosta_Err
        

        'Ladder 10 - 2 - 2010
        'Chequea si hay costa en los tiles proximos al usuario
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If ((MapData(Map, X, Y).Graphic(1) >= 22552 And MapData(Map, X, Y).Graphic(1) <= 22599) Or (MapData(Map, X, Y).Graphic(1) >= 7283 And MapData(Map, X, Y).Graphic(1) <= 7378) Or (MapData(Map, X, Y).Graphic(1) >= 13387 And MapData(Map, X, Y).Graphic(1) <= 13482)) And MapData(Map, X, Y).Graphic(2) = 0 Then
104             HayCosta = True
            Else
106             HayCosta = False

            End If

        Else
108         HayCosta = False

        End If

        
        Exit Function

HayCosta_Err:
        Call RegistrarError(Err.Number, Err.description, "General.HayCosta", Erl)
        Resume Next
        
End Function

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayAgua_Err
        

100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If ((MapData(Map, X, Y).Graphic(1) >= 1505 And MapData(Map, X, Y).Graphic(1) <= 1520) Or (MapData(Map, X, Y).Graphic(1) >= 24223 And MapData(Map, X, Y).Graphic(1) <= 24238) Or (MapData(Map, X, Y).Graphic(1) >= 24303 And MapData(Map, X, Y).Graphic(1) <= 24318) Or (MapData(Map, X, Y).Graphic(1) >= 468 And MapData(Map, X, Y).Graphic(1) <= 483) Or (MapData(Map, X, Y).Graphic(1) >= 44668 And MapData(Map, X, Y).Graphic(1) <= 44939) Or (MapData(Map, X, Y).Graphic(1) >= 24143 And MapData(Map, X, Y).Graphic(1) <= 24158)) And MapData(Map, X, Y).Graphic(2) = 0 Then
104             HayAgua = True
            Else
106             HayAgua = False

            End If

        Else
108         HayAgua = False

        End If

        
        Exit Function

HayAgua_Err:
        Call RegistrarError(Err.Number, Err.description, "General.HayAgua", Erl)
        Resume Next
        
End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayLava_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        '***************************************************
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
104             HayLava = True
            Else
106             HayLava = False

            End If

        Else
108         HayLava = False

        End If

        
        Exit Function

HayLava_Err:
        Call RegistrarError(Err.Number, Err.description, "General.HayLava", Erl)
        Resume Next
        
End Function

Sub ApagarFogatas()

    'Ladder /ApagarFogatas
    On Error GoTo errHandler

    Dim obj As obj
        obj.ObjIndex = FOGATA_APAG
        obj.Amount = 1

    Dim MapaActual As Long
    Dim Y          As Long
    Dim X          As Long

    For MapaActual = 1 To NumMaps
        For Y = YMinMapSize To YMaxMapSize
            For X = XMinMapSize To XMaxMapSize

                If MapInfo(MapaActual).lluvia Then
                
                    If MapData(MapaActual, X, Y).ObjInfo.ObjIndex = FOGATA Then
                    
                        Call EraseObj(MAX_INVENTORY_OBJS, MapaActual, X, Y)
                        Call MakeObj(obj, MapaActual, X, Y)

                    End If

                End If

            Next X
        Next Y
    Next MapaActual

    Exit Sub
    
errHandler:
    Call LogError("Error producido al apagar las fogatas de " & X & "-" & Y & " del mapa: " & MapaActual & "    -" & Err.description)

End Sub

Sub EnviarSpawnList(ByVal Userindex As Integer)
        
        On Error GoTo EnviarSpawnList_Err
        

        Dim K          As Long
        Dim npcNames() As String

100     Debug.Print UBound(SpawnList)
102     ReDim npcNames(1 To UBound(SpawnList)) As String

104     For K = 1 To UBound(SpawnList)
106         npcNames(K) = SpawnList(K).NpcName
108     Next K

110     Call WriteSpawnList(Userindex, npcNames())

        
        Exit Sub

EnviarSpawnList_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EnviarSpawnList", Erl)
        Resume Next
        
End Sub

Sub ConfigListeningSocket(ByRef obj As Object, ByVal Port As Integer)
        
        On Error GoTo ConfigListeningSocket_Err
        
        #If UsarQueSocket = 0 Then

100         obj.AddressFamily = AF_INET
102         obj.Protocol = IPPROTO_IP
104         obj.SocketType = SOCK_STREAM
106         obj.Binary = False
108         obj.Blocking = False
110         obj.BufferSize = 1024
112         obj.LocalPort = Port
114         obj.backlog = 5
116         obj.listen

        #End If

        
        Exit Sub

ConfigListeningSocket_Err:
        Call RegistrarError(Err.Number, Err.description, "General.ConfigListeningSocket", Erl)
        Resume Next
        
End Sub

Public Sub LeerLineaComandos()
        
        On Error GoTo LeerLineaComandos_Err
        

        Dim rdata As String

100     rdata = Command
102     rdata = Right$(rdata, Len(rdata))
104     ClaveApertura = ReadField(1, rdata, Asc("*")) ' NICK

        
        Exit Sub

LeerLineaComandos_Err:
        Call RegistrarError(Err.Number, Err.description, "General.LeerLineaComandos", Erl)
        Resume Next
        
End Sub

Sub Main()

    On Error Resume Next

    Call LeerLineaComandos
    
    CargarRanking
    
    Dim f    As Date

    Dim abro As Boolean
    
    ChDir App.Path
    ChDrive App.Path
    
    abro = True
    Prision.Map = 23
    Libertad.Map = 23
    
    Prision.X = 72
    Prision.Y = 52
    Libertad.X = 73
    Libertad.Y = 73
    
    LastBackup = Format(Now, "Short Time")
    minutos = Format(Now, "Short Time")
    
    IniPath = App.Path & "\"

    LevelSkill(1).LevelValue = 3
    LevelSkill(2).LevelValue = 5
    LevelSkill(3).LevelValue = 7
    LevelSkill(4).LevelValue = 10
    LevelSkill(5).LevelValue = 13
    LevelSkill(6).LevelValue = 15
    LevelSkill(7).LevelValue = 17
    LevelSkill(8).LevelValue = 20
    LevelSkill(9).LevelValue = 23
    LevelSkill(10).LevelValue = 25
    LevelSkill(11).LevelValue = 27
    LevelSkill(12).LevelValue = 30
    LevelSkill(13).LevelValue = 33
    LevelSkill(14).LevelValue = 35
    LevelSkill(15).LevelValue = 37
    LevelSkill(16).LevelValue = 40
    LevelSkill(17).LevelValue = 43
    LevelSkill(18).LevelValue = 45
    LevelSkill(19).LevelValue = 47
    LevelSkill(20).LevelValue = 50
    LevelSkill(21).LevelValue = 53
    LevelSkill(22).LevelValue = 55
    LevelSkill(23).LevelValue = 57
    LevelSkill(24).LevelValue = 60
    LevelSkill(25).LevelValue = 63
    LevelSkill(26).LevelValue = 65
    LevelSkill(27).LevelValue = 67
    LevelSkill(28).LevelValue = 70
    LevelSkill(29).LevelValue = 73
    LevelSkill(30).LevelValue = 75
    LevelSkill(31).LevelValue = 77
    LevelSkill(32).LevelValue = 80
    LevelSkill(33).LevelValue = 83
    LevelSkill(34).LevelValue = 85
    LevelSkill(35).LevelValue = 87
    LevelSkill(36).LevelValue = 90
    LevelSkill(37).LevelValue = 93
    LevelSkill(38).LevelValue = 95
    LevelSkill(39).LevelValue = 97
    LevelSkill(40).LevelValue = 100
    LevelSkill(41).LevelValue = 100
    LevelSkill(42).LevelValue = 100
    LevelSkill(43).LevelValue = 100
    LevelSkill(44).LevelValue = 100
    LevelSkill(45).LevelValue = 100
    LevelSkill(46).LevelValue = 100
    LevelSkill(47).LevelValue = 100
    LevelSkill(48).LevelValue = 100
    LevelSkill(49).LevelValue = 100
    LevelSkill(50).LevelValue = 100
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.Drow) = "Elfo Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    'ListaRazas(eRaza.Orco) = "Orco"
    
    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Trabajador) = "Trabajador"
    
    SkillsNames(eSkill.magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Wrestling) = "Artes Marciales"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.equitacion) = "Equitación"
    SkillsNames(eSkill.Resistencia) = "Resistencia Mágica"

    SkillsNames(eSkill.Talar) = "Tala"
    SkillsNames(eSkill.Pescar) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Alquimia) = "Alquimia"
    SkillsNames(eSkill.Sastreria) = "Sastreria"
   
    ListaAtributos(eAtributos.Fuerza) = "Fuerza"
    ListaAtributos(eAtributos.Agilidad) = "Agilidad"
    ListaAtributos(eAtributos.Inteligencia) = "Inteligencia"
    ListaAtributos(eAtributos.Constitucion) = "Constitucion"
    ListaAtributos(eAtributos.Carisma) = "Carisma"
    
    centinelaActivado = False
    
    frmCargando.Show
    
    Call InitTesoro
    Call InitRegalo
    
    'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")
    
    frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    IniPath = App.Path & "\"
    
    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    DoEvents
    
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    
    Call LoadGuildsDB
    
    Call LoadConfiguraciones
    Call CargarEventos
    Call CargarCodigosDonador
    Call loadAdministrativeUsers

    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
    MaxUsers = 0
    Call LoadSini
    Call LoadIntervalos
    Call CargarForbidenWords
    Call CargaApuestas
    Call CargarSpawnList
    Call LoadMotd
    Call BanIpCargar

    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    '*************************************************
    
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    'Call LoadOBJData
    Call LoadOBJData
        
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
        
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
    Call LoadObjCarpintero
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
    Call LoadObjAlquimista
    
    frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"
    Call LoadObjSastre
    
    frmCargando.Label1(2).Caption = "Cargando Pesca"
    Call LoadPesca
    
    frmCargando.Label1(2).Caption = "Cargando Recursos Especiales"
    Call LoadRecursosEspeciales
    
    frmCargando.Label1(2).Caption = "Cargando Balance.Dat"
    Call LoadBalance    '4/01/08 Pablo ToxicWaste
    
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando BackUp"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    
    ' Pretorianos
    frmCargando.Label1(2).Caption = "Cargando Pretorianos.dat"
    Call LoadPretorianData
    
    frmCargando.Label1(2).Caption = "Cargando Logros.ini"
    Call CargarLogros ' Ladder 22/04/2015
    
    frmCargando.Label1(2).Caption = "Cargando Baneos Temporales"
    Call LoadBans
    
    frmCargando.Label1(2).Caption = "Cargando Usuarios Donadores"
    Call LoadDonadores
    Call LoadObjDonador
    Call LoadQuests

    EstadoGlobal = True
    
    Call InicializarLimpieza

    'Comentado porque hay worldsave en ese mapa!
    'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Dim LoopC As Integer
    
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    With frmMain
        .AutoSave.Enabled = True
        '.tLluvia.Enabled = True
        '.tPiqueteC.Enabled = True
        .GameTimer.Enabled = True
        .Auditoria.Enabled = True
        .KillLog.Enabled = True
        .TIMER_AI.Enabled = True

        '.npcataca.Enabled = True
    End With
    
    Subasta.SubastaHabilitada = True
    Subasta.HaySubastaActiva = False
    Call ResetMeteo
    
    frmCargando.Label1(2).Caption = "Conectando base de datos y limpiando usuarios logueados"
    
    If Database_Enabled Then
        'Conecto base de datos
        Call Database_Connect
        
        'Reinicio los users online
        Call SetUsersLoggedDatabase(0)
        
        'Leo el record de usuarios
        RecordUsuarios = LeerRecordUsuariosDatabase()
        
        'Tarea pesada
        Call LogoutAllUsersAndAccounts
    End If
    
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    'Configuracion de los sockets
    
    Call SecurityIp.InitIpTables(1000)
    
    #If UsarQueSocket = 1 Then
    
        If LastSockListen >= 0 Then Call apiclosesocket(LastSockListen) 'Cierra el socket de escucha
        Call IniciaWsApi(frmMain.hWnd)
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

        If SockListen <> -1 Then
            Call WriteVar(IniPath & "Server.ini", "INIT", "LastSockListen", SockListen) ' Guarda el socket escuchando
        Else
            MsgBox "Ha ocurrido un error al iniciar el socket del Servidor.", vbCritical + vbOKOnly

        End If
    
    #ElseIf UsarQueSocket = 0 Then
    
        frmCargando.Label1(2).Caption = "Configurando Sockets"
    
        frmMain.Socket2(0).AddressFamily = AF_INET
        frmMain.Socket2(0).Protocol = IPPROTO_IP
        frmMain.Socket2(0).SocketType = SOCK_STREAM
        frmMain.Socket2(0).Binary = False
        frmMain.Socket2(0).Blocking = False
        frmMain.Socket2(0).BufferSize = 2048
    
        Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
    #ElseIf UsarQueSocket = 2 Then
    
        frmMain.Serv.Iniciar Puerto
    
    #ElseIf UsarQueSocket = 3 Then
    
        frmMain.TCPServ.Encolar True
        frmMain.TCPServ.IniciarTabla 1009
        frmMain.TCPServ.SetQueueLim 51200
        frmMain.TCPServ.Iniciar Puerto
    
    #End If
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
    Call GetHoraActual
    
    HoraFanstasia = 720
    Unload frmCargando
    
    'Log
    Dim n As Integer

    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
    Close #n
    
    'Ocultar
    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)
    End If
    
    tInicioServer = GetTickCount() And &H7FFFFFFF
    'Call InicializaEstadisticas

End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
        '*****************************************************************
        'Se fija si existe el archivo
        '*****************************************************************
        
        On Error GoTo FileExist_Err
        
100     FileExist = LenB(dir$(File, FileType)) <> 0

        
        Exit Function

FileExist_Err:
        Call RegistrarError(Err.Number, Err.description, "General.FileExist", Erl)
        Resume Next
        
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
        
        On Error GoTo ReadField_Err
        

        '*****************************************************************
        'Gets a field from a string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        'Gets a field from a delimited string
        '*****************************************************************
        Dim i          As Long

        Dim LastPos    As Long

        Dim CurrentPos As Long

        Dim delimiter  As String * 1
    
100     delimiter = Chr$(SepASCII)
    
102     For i = 1 To Pos
104         LastPos = CurrentPos
106         CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i
    
110     If CurrentPos = 0 Then
112         ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
        Else
114         ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

        End If

        
        Exit Function

ReadField_Err:
        Call RegistrarError(Err.Number, Err.description, "General.ReadField", Erl)
        Resume Next
        
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
        
        On Error GoTo MapaValido_Err
        
100     MapaValido = Map >= 1 And Map <= NumMaps

        
        Exit Function

MapaValido_Err:
        Call RegistrarError(Err.Number, Err.description, "General.MapaValido", Erl)
        Resume Next
        
End Function

Sub MostrarNumUsers()
        
        On Error GoTo MostrarNumUsers_Err
        

100     Call SendData(SendTarget.ToAll, 0, PrepareMessageOnlineUser(NumUsers))
102     frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
    
104     Call SetUsersLoggedDatabase(NumUsers)

        
        Exit Sub

MostrarNumUsers_Err:
        Call RegistrarError(Err.Number, Err.description, "General.MostrarNumUsers", Erl)
        Resume Next
        
End Sub

Public Sub LogCriticEvent(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogError(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogConsulta(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\ConsultasGM.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogStatic(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogTarea(Desc As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogClanes(ByVal str As String)
        
        On Error GoTo LogClanes_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogClanes_Err:
        Call RegistrarError(Err.Number, Err.description, "General.LogClanes", Erl)
        Resume Next
        
End Sub

Public Sub LogIP(ByVal str As String)
        
        On Error GoTo LogIP_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\IP.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogIP_Err:
        Call RegistrarError(Err.Number, Err.description, "General.LogIP", Erl)
        Resume Next
        
End Sub

Public Sub LogDesarrollo(ByVal str As String)
        
        On Error GoTo LogDesarrollo_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogDesarrollo_Err:
        Call RegistrarError(Err.Number, Err.description, "General.LogDesarrollo", Erl)
        Resume Next
        
End Sub

Public Sub LogGM(nombre As String, texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogDatabaseError(Desc As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Database.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " - " & Desc
    Close #nfile
    
    Exit Sub
    
    Debug.Print Desc
    
errHandler:

End Sub

Public Sub SaveDayStats()
    ''On Error GoTo errhandler
    ''
    ''Dim nfile As Integer
    ''nfile = FreeFile ' obtenemos un canal
    ''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
    ''
    ''Print #nfile, "<stats>"
    ''Print #nfile, "<ao>"
    ''Print #nfile, "<dia>" & Date & "</dia>"
    ''Print #nfile, "<hora>" & Time & "</hora>"
    ''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
    ''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
    ''Print #nfile, "</ao>"
    ''Print #nfile, "</stats>"
    ''
    ''
    ''Close #nfile
    
    On Error GoTo SaveDayStats_Err
    
    Exit Sub

errHandler:

    
    Exit Sub

SaveDayStats_Err:
    Call RegistrarError(Err.Number, Err.description, "General.SaveDayStats", Erl)
    Resume Next
    
End Sub

Public Sub LogAsesinato(texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal

    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal

    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogHackAttemp(texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogCheating(texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogAntiCheat(texto As String)

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, ""
    Close #nfile

    Exit Sub

errHandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
        
        On Error GoTo ValidInputNP_Err
        

        Dim Arg As String

        Dim i   As Integer

100     For i = 1 To 33

102         Arg = ReadField(i, cad, 44)

104         If LenB(Arg) = 0 Then Exit Function

106     Next i

108     ValidInputNP = True

        
        Exit Function

ValidInputNP_Err:
        Call RegistrarError(Err.Number, Err.description, "General.ValidInputNP", Erl)
        Resume Next
        
End Function

Sub Restart()

    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    On Error Resume Next

    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

    Dim LoopC As Long
  
    #If UsarQueSocket = 0 Then

        frmMain.Socket1.Cleanup
        frmMain.Socket1.Startup
      
        frmMain.Socket2(0).Cleanup
        frmMain.Socket2(0).Startup

    #ElseIf UsarQueSocket = 1 Then

        'Cierra el socket de escucha
        If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
        'Inicia el socket de escucha
        SockListen = ListenForConnect(Puerto, hWndMsg, "")

    #ElseIf UsarQueSocket = 2 Then

    #End If

    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next

    'Initialize statistics!!
    'Call Statistics.Initialize

    For LoopC = 1 To UBound(UserList())
        Set UserList(LoopC).incomingData = Nothing
        Set UserList(LoopC).outgoingData = Nothing
    Next LoopC

    ReDim UserList(1 To MaxUsers) As user

    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnID = -1
        UserList(LoopC).ConnIDValida = False
        Set UserList(LoopC).incomingData = New clsByteQueue
        Set UserList(LoopC).outgoingData = New clsByteQueue
    Next LoopC

    LastUser = 0
    NumUsers = 0

    Call FreeNPCs
    Call FreeCharIndexes

    Call LoadSini
    Call LoadIntervalos
    Call LoadOBJData
    Call LoadPesca
    Call LoadRecursosEspeciales

    Call LoadMapData

    Call CargarHechizos

    #If UsarQueSocket = 0 Then

        '*****************Setup socket
        frmMain.Socket1.AddressFamily = AF_INET
        frmMain.Socket1.Protocol = IPPROTO_IP
        frmMain.Socket1.SocketType = SOCK_STREAM
        frmMain.Socket1.Binary = False
        frmMain.Socket1.Blocking = False
        frmMain.Socket1.BufferSize = 1024

        frmMain.Socket2(0).AddressFamily = AF_INET
        frmMain.Socket2(0).Protocol = IPPROTO_IP
        frmMain.Socket2(0).SocketType = SOCK_STREAM
        frmMain.Socket2(0).Blocking = False
        frmMain.Socket2(0).BufferSize = 2048

        'Escucha
        frmMain.Socket1.LocalPort = val(Puerto)
        frmMain.Socket1.listen

    #ElseIf UsarQueSocket = 1 Then

    #ElseIf UsarQueSocket = 2 Then

    #End If

    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

    'Log it
    Dim n As Integer

    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & Time & " servidor reiniciado."
    Close #n

    'Ocultar

    If HideMe = 1 Then
        Call frmMain.InitMain(1)
    Else
        Call frmMain.InitMain(0)

    End If
  
End Sub

Public Function Intemperie(ByVal Userindex As Integer) As Boolean
        
        On Error GoTo Intemperie_Err
        
    
100     If MapInfo(UserList(Userindex).Pos.Map).zone <> "DUNGEON" Then
102         If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger <> 1 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger <> 2 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger < 10 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger <> 4 Then Intemperie = True
        Else
104         Intemperie = False

        End If
    
        
        Exit Function

Intemperie_Err:
        Call RegistrarError(Err.Number, Err.description, "General.Intemperie", Erl)
        Resume Next
        
End Function

Public Sub EfectoFrio(ByVal Userindex As Integer)
        
        On Error GoTo EfectoFrio_Err
        
        If Not Intemperie(Userindex) Then Exit Sub
        
        Dim modifi As Integer
        
100     With UserList(Userindex)
            
            If .flags.Desnudo = 0 Then Exit Sub
            
102         If .Counters.Frio < IntervaloFrio Then
104             .Counters.Frio = .Counters.Frio + 1

            Else

106             If MapInfo(.Pos.Map).terrain = Nieve Then
108                 Call WriteConsoleMsg(Userindex, "¡¡Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)

110                 modifi = Porcentaje(.Stats.MaxHp, 5)

112                 .Stats.MinHp = .Stats.MinHp - modifi
            
114                 If .Stats.MinHp < 1 Then

116                     Call WriteConsoleMsg(Userindex, "¡¡Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)

118                     .Stats.MinHp = 0

120                     Call UserDie(Userindex)

                    End If
            
122                 Call WriteUpdateHP(Userindex)
                End If
        
128             .Counters.Frio = 0

            End If
        
        End With
        
        Exit Sub

EfectoFrio_Err:
130     Call RegistrarError(Err.Number, Err.description, "General.EfectoFrio", Erl)

132     Resume Next
        
End Sub

Public Sub EfectoLava(ByVal Userindex As Integer)
        
        On Error GoTo EfectoLava_Err

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        'If user is standing on lava, take health points from him
        '***************************************************
        
100     With UserList(Userindex)
        
102         If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
104             .Counters.Lava = .Counters.Lava + 1
        
            Else

106             If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
108                 Call WriteConsoleMsg(Userindex, "¡¡Quitate de la lava, te estás quemando!!.", FontTypeNames.FONTTYPE_INFO)
110                 .Stats.MinHp = .Stats.MinHp - Porcentaje(.Stats.MaxHp, 5)
            
112                 If .Stats.MinHp < 1 Then
114                     Call WriteConsoleMsg(Userindex, "¡¡Has muerto quemado!!.", FontTypeNames.FONTTYPE_INFO)
116                     .Stats.MinHp = 0
118                     Call UserDie(Userindex)

                    End If
            
120                 Call WriteUpdateHP(Userindex)

                End If
        
122             .Counters.Lava = 0

            End If
        
        End With
        

        
        Exit Sub

EfectoLava_Err:
124     Call RegistrarError(Err.Number, Err.description, "General.EfectoLava", Erl)

126     Resume Next
        
End Sub

Public Sub EfectoInvisibilidad(ByVal Userindex As Integer)
        
        On Error GoTo EfectoInvisibilidad_Err
        

100     If UserList(Userindex).Counters.Invisibilidad > 0 Then
102         UserList(Userindex).Counters.Invisibilidad = UserList(Userindex).Counters.Invisibilidad - 1
        Else
104         UserList(Userindex).Counters.Invisibilidad = 0
106         UserList(Userindex).flags.invisible = 0

108         If UserList(Userindex).flags.Oculto = 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
110             Call WriteLocaleMsg(Userindex, "307", FontTypeNames.FONTTYPE_INFO)
112             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(Userindex).Char.CharIndex, False))
114             Call WriteContadores(Userindex)

            End If

        End If

        
        Exit Sub

EfectoInvisibilidad_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoInvisibilidad", Erl)
        Resume Next
        
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
        
        On Error GoTo EfectoParalisisNpc_Err
        

100     If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
102         Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
        Else
104         Npclist(NpcIndex).flags.Paralizado = 0
106         Npclist(NpcIndex).flags.Inmovilizado = 0

        End If

        
        Exit Sub

EfectoParalisisNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoParalisisNpc", Erl)
        Resume Next
        
End Sub

Public Sub EfectoCegueEstu(ByVal Userindex As Integer)
        
        On Error GoTo EfectoCegueEstu_Err
        

100     If UserList(Userindex).Counters.Ceguera > 0 Then
102         UserList(Userindex).Counters.Ceguera = UserList(Userindex).Counters.Ceguera - 1
        Else

104         If UserList(Userindex).flags.Ceguera = 1 Then
106             UserList(Userindex).flags.Ceguera = 0
108             Call WriteBlindNoMore(Userindex)

            End If

        End If

        
        Exit Sub

EfectoCegueEstu_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoCegueEstu", Erl)
        Resume Next
        
End Sub

Public Sub EfectoEstupidez(ByVal Userindex As Integer)
        
        On Error GoTo EfectoEstupidez_Err
        

100     If UserList(Userindex).Counters.Estupidez > 0 Then
102         UserList(Userindex).Counters.Estupidez = UserList(Userindex).Counters.Estupidez - 1

        Else

104         If UserList(Userindex).flags.Estupidez = 1 Then
106             UserList(Userindex).flags.Estupidez = 0
108             Call WriteDumbNoMore(Userindex)

            End If

        End If

        
        Exit Sub

EfectoEstupidez_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoEstupidez", Erl)
        Resume Next
        
End Sub

Public Sub EfectoParalisisUser(ByVal Userindex As Integer)
        
        On Error GoTo EfectoParalisisUser_Err
        

100     If UserList(Userindex).Counters.Paralisis > 0 Then
102         UserList(Userindex).Counters.Paralisis = UserList(Userindex).Counters.Paralisis - 1
        Else
104         UserList(Userindex).flags.Paralizado = 0
            'UserList(UserIndex).Flags.AdministrativeParalisis = 0
106         Call WriteParalizeOK(Userindex)

        End If

        
        Exit Sub

EfectoParalisisUser_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoParalisisUser", Erl)
        Resume Next
        
End Sub

Public Sub EfectoVelocidadUser(ByVal Userindex As Integer)
        
        On Error GoTo EfectoVelocidadUser_Err
        

100     If UserList(Userindex).Counters.Velocidad > 0 Then
102         UserList(Userindex).Counters.Velocidad = UserList(Userindex).Counters.Velocidad - 1
        Else
104         UserList(Userindex).Char.speeding = UserList(Userindex).flags.VelocidadBackup
    
            'Call WriteVelocidadToggle(UserIndex)
106         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(UserList(Userindex).Char.CharIndex, UserList(Userindex).flags.VelocidadBackup))
108         UserList(Userindex).flags.VelocidadBackup = 0

        End If

        
        Exit Sub

EfectoVelocidadUser_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoVelocidadUser", Erl)
        Resume Next
        
End Sub

Public Sub EfectoMaldicionUser(ByVal Userindex As Integer)
        
        On Error GoTo EfectoMaldicionUser_Err
        

100     If UserList(Userindex).Counters.Maldicion > 0 Then
102         UserList(Userindex).Counters.Maldicion = UserList(Userindex).Counters.Maldicion - 1
    
        Else
104         UserList(Userindex).flags.Maldicion = 0
106         Call WriteConsoleMsg(Userindex, "¡La magia perdió su efecto! Ya podes atacar.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)

            'Call WriteParalizeOK(UserIndex)
        End If

        
        Exit Sub

EfectoMaldicionUser_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoMaldicionUser", Erl)
        Resume Next
        
End Sub

Public Sub EfectoInmoUser(ByVal Userindex As Integer)
        
        On Error GoTo EfectoInmoUser_Err
        

100     If UserList(Userindex).Counters.Inmovilizado > 0 Then
102         UserList(Userindex).Counters.Inmovilizado = UserList(Userindex).Counters.Inmovilizado - 1
        Else
104         UserList(Userindex).flags.Inmovilizado = 0
            'UserList(UserIndex).Flags.AdministrativeParalisis = 0
106         Call WriteInmovilizaOK(Userindex)

        End If

        
        Exit Sub

EfectoInmoUser_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoInmoUser", Erl)
        Resume Next
        
End Sub

Public Sub RecStamina(ByVal Userindex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        
        On Error GoTo RecStamina_Err
        

100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 1 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 2 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 4 Then Exit Sub

        Dim massta As Integer

102     If UserList(Userindex).Stats.MinSta < UserList(Userindex).Stats.MaxSta Then

104         If UserList(Userindex).Counters.STACounter < Intervalo Then
106             UserList(Userindex).Counters.STACounter = UserList(Userindex).Counters.STACounter + 1
            Else
        
108             UserList(Userindex).Counters.STACounter = 0

110             If UserList(Userindex).flags.Desnudo And Not UserList(Userindex).flags.Montado Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
112             If UserList(Userindex).Counters.Trabajando > 0 Then Exit Sub  'Trabajando no sube energía. (ToxicWaste)
         
                ' If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub 'Ladder, se ve que esta linea la agregue yo, pero no sirve.

114             EnviarStats = True
        
                Dim Suerte As Integer

116             If UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 10 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= -1 Then
118                 Suerte = 5
120             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 20 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 11 Then
122                 Suerte = 7
124             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 30 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 21 Then
126                 Suerte = 9
128             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 40 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 31 Then
130                 Suerte = 11
132             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 50 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 41 Then
134                 Suerte = 13
136             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 60 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 51 Then
138                 Suerte = 15
140             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 70 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 61 Then
142                 Suerte = 17
144             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 80 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 71 Then
146                 Suerte = 19
148             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) <= 90 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 81 Then
150                 Suerte = 21
152             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) < 100 And UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) >= 91 Then
154                 Suerte = 23
156             ElseIf UserList(Userindex).Stats.UserSkills(eSkill.Supervivencia) = 100 Then
158                 Suerte = 25

                End If
        
160             If UserList(Userindex).flags.RegeneracionSta = 1 Then
162                 Suerte = 45

                End If
        
164             massta = RandomNumber(1, Porcentaje(UserList(Userindex).Stats.MaxSta, Suerte))
166             UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta + massta

168             If UserList(Userindex).Stats.MinSta > UserList(Userindex).Stats.MaxSta Then
170                 UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MaxSta

                End If

            End If

        End If

        
        Exit Sub

RecStamina_Err:
        Call RegistrarError(Err.Number, Err.description, "General.RecStamina", Erl)
        Resume Next
        
End Sub

Public Sub EfectoVeneno(ByVal Userindex As Integer)
        
        On Error GoTo EfectoVeneno_Err
        

        Dim n As Integer

100     If UserList(Userindex).Counters.Veneno < IntervaloVeneno Then
102         UserList(Userindex).Counters.Veneno = UserList(Userindex).Counters.Veneno + 1
        Else
            'Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas moriras.", FontTypeNames.FONTTYPE_VENENO)
104         Call WriteLocaleMsg(Userindex, "47", FontTypeNames.FONTTYPE_VENENO)
106         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Envenena, 30, False))
108         UserList(Userindex).Counters.Veneno = 0
110         n = RandomNumber(3, 6)
112         n = n * UserList(Userindex).flags.Envenenado
114         UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - n

116         If UserList(Userindex).Stats.MinHp < 1 Then Call UserDie(Userindex)
118         Call WriteUpdateHP(Userindex)

        End If

        
        Exit Sub

EfectoVeneno_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoVeneno", Erl)
        Resume Next
        
End Sub

Public Sub EfectoAhogo(ByVal Userindex As Integer)
        
        On Error GoTo EfectoAhogo_Err
        

        Dim n As Integer

100     If RequiereOxigeno(UserList(Userindex).Pos.Map) Then
102         If UserList(Userindex).Counters.Ahogo < 70 Then
104             UserList(Userindex).Counters.Ahogo = UserList(Userindex).Counters.Ahogo + 1
            Else
106             Call WriteConsoleMsg(Userindex, "Te estas ahogando.. si no consigues oxigeno moriras.", FontTypeNames.FONTTYPE_EJECUCION)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 205, 30, False))
108             UserList(Userindex).Counters.Ahogo = 0
110             n = RandomNumber(150, 200)
112             UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - n

114             If UserList(Userindex).Stats.MinHp < 1 Then
116                 Call UserDie(Userindex)
118                 UserList(Userindex).flags.Ahogandose = 0

                End If

120             Call WriteUpdateHP(Userindex)

            End If

        Else
122         UserList(Userindex).flags.Ahogandose = 0

        End If

        
        Exit Sub

EfectoAhogo_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoAhogo", Erl)
        Resume Next
        
End Sub

Public Sub EfectoIncineramiento(ByVal Userindex As Integer, ByRef EnviarStats As Boolean)
        
        On Error GoTo EfectoIncineramiento_Err
        

        Dim n As Integer
 
100     If UserList(Userindex).Counters.Incineracion < IntervaloIncineracion Then
102         UserList(Userindex).Counters.Incineracion = UserList(Userindex).Counters.Incineracion + 1
        Else
104         Call WriteConsoleMsg(Userindex, "Te estas incinerando,si no te curas moriras.", FontTypeNames.FONTTYPE_INFO)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Incinerar, 30, False))
106         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 73, 0))
108         UserList(Userindex).Counters.Incineracion = 0
110         n = RandomNumber(40, 80)
112         UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - n

114         If UserList(Userindex).Stats.MinHp < 1 Then Call UserDie(Userindex)
116         Call WriteUpdateHP(Userindex)

        End If
 
        
        Exit Sub

EfectoIncineramiento_Err:
        Call RegistrarError(Err.Number, Err.description, "General.EfectoIncineramiento", Erl)
        Resume Next
        
End Sub

Public Sub DuracionPociones(ByVal Userindex As Integer)
        
        On Error GoTo DuracionPociones_Err
        

        'Controla la duracion de las pociones
100     If UserList(Userindex).flags.DuracionEfecto > 0 Then
102         UserList(Userindex).flags.DuracionEfecto = UserList(Userindex).flags.DuracionEfecto - 1

104         If UserList(Userindex).flags.DuracionEfecto = 0 Then
106             UserList(Userindex).flags.TomoPocion = False
108             UserList(Userindex).flags.TipoPocion = 0

                'volvemos los atributos al estado normal
                Dim loopX As Integer

110             For loopX = 1 To NUMATRIBUTOS
112                 UserList(Userindex).Stats.UserAtributos(loopX) = UserList(Userindex).Stats.UserAtributosBackUP(loopX)
                Next
114             Call WriteFYA(Userindex)

            End If

        End If

        
        Exit Sub

DuracionPociones_Err:
        Call RegistrarError(Err.Number, Err.description, "General.DuracionPociones", Erl)
        Resume Next
        
End Sub

Public Sub HambreYSed(ByVal Userindex As Integer, ByRef fenviarAyS As Boolean)
        
        On Error GoTo HambreYSed_Err
        

100     If Not UserList(Userindex).flags.Privilegios And PlayerType.user Then Exit Sub
102     If UserList(Userindex).flags.BattleModo = 1 Then Exit Sub

        'Sed
104     If UserList(Userindex).Stats.MinAGU > 0 Then
106         If UserList(Userindex).Counters.AGUACounter < IntervaloSed Then
108             UserList(Userindex).Counters.AGUACounter = UserList(Userindex).Counters.AGUACounter + 1
            Else
110             UserList(Userindex).Counters.AGUACounter = 0
112             UserList(Userindex).Stats.MinAGU = UserList(Userindex).Stats.MinAGU - 10
        
114             If UserList(Userindex).Stats.MinAGU <= 0 Then
116                 UserList(Userindex).Stats.MinAGU = 0
118                 UserList(Userindex).flags.Sed = 1

                End If
        
120             fenviarAyS = True

            End If

        End If

        'hambre
122     If UserList(Userindex).Stats.MinHam > 0 Then
124         If UserList(Userindex).Counters.COMCounter < IntervaloHambre Then
126             UserList(Userindex).Counters.COMCounter = UserList(Userindex).Counters.COMCounter + 1
            Else
128             UserList(Userindex).Counters.COMCounter = 0
130             UserList(Userindex).Stats.MinHam = UserList(Userindex).Stats.MinHam - 10

132             If UserList(Userindex).Stats.MinHam <= 0 Then
134                 UserList(Userindex).Stats.MinHam = 0
136                 UserList(Userindex).flags.Hambre = 1

                End If

138             fenviarAyS = True

            End If

        End If

        
        Exit Sub

HambreYSed_Err:
        Call RegistrarError(Err.Number, Err.description, "General.HambreYSed", Erl)
        Resume Next
        
End Sub

Public Sub Sanar(ByVal Userindex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        
        On Error GoTo Sanar_Err
        

100     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 1 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 2 And MapData(UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y).trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
102     If UserList(Userindex).Stats.MinHp < UserList(Userindex).Stats.MaxHp Then
104         If UserList(Userindex).flags.RegeneracionHP = 1 Then
106             Intervalo = 400

            End If
    
108         If UserList(Userindex).Counters.HPCounter < Intervalo Then
110             UserList(Userindex).Counters.HPCounter = UserList(Userindex).Counters.HPCounter + 1
            Else
112             mashit = RandomNumber(2, Porcentaje(UserList(Userindex).Stats.MaxSta, 5))
        
114             UserList(Userindex).Counters.HPCounter = 0
116             UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp + mashit

118             If UserList(Userindex).Stats.MinHp > UserList(Userindex).Stats.MaxHp Then UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
120             Call WriteConsoleMsg(Userindex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
122             EnviarStats = True

            End If

        End If

        
        Exit Sub

Sanar_Err:
        Call RegistrarError(Err.Number, Err.description, "General.Sanar", Erl)
        Resume Next
        
End Sub

Public Sub CargaNpcsDat()
        
        On Error GoTo CargaNpcsDat_Err
        

        Dim npcfile As String
    
100     npcfile = DatPath & "NPCs.dat"
102     Call LeerNPCs.Initialize(npcfile)
    
        'npcfile = DatPath & "NPCs-HOSTILES.dat"
        'Call LeerNPCsHostiles.Initialize(npcfile)
        
        Exit Sub

CargaNpcsDat_Err:
        Call RegistrarError(Err.Number, Err.description, "General.CargaNpcsDat", Erl)
        Resume Next
        
End Sub

Sub PasarSegundo()

    On Error GoTo errHandler

    Dim i    As Long

    Dim h    As Byte

    Dim Mapa As Integer

    Dim X    As Byte

    Dim Y    As Byte
    
    If CuentaRegresivaTimer > 0 Then
        If CuentaRegresivaTimer > 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(CuentaRegresivaTimer - 1 & " segundos...!", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ya!!!", FontTypeNames.FONTTYPE_FIGHT))

        End If

        CuentaRegresivaTimer = CuentaRegresivaTimer - 1

    End If
    
    For i = 1 To LastUser

        If UserList(i).flags.Silenciado = 1 Then
            UserList(i).flags.SegundosPasados = UserList(i).flags.SegundosPasados + 1

            If UserList(i).flags.SegundosPasados = 60 Then
                UserList(i).flags.MinutosRestantes = UserList(i).flags.MinutosRestantes - 1
                UserList(i).flags.SegundosPasados = 0

            End If
            
            If UserList(i).flags.MinutosRestantes = 0 Then
                UserList(i).flags.SegundosPasados = 0
                UserList(i).flags.Silenciado = 0
                UserList(i).flags.MinutosRestantes = 0
                Call WriteConsoleMsg(i, "Has sido liberado del silencio.", FontTypeNames.FONTTYPE_SERVER)

            End If

        End If

        With UserList(i)
        
            If .flags.invisible = 1 Then Call EfectoInvisibilidad(i)
            If .flags.BattleModo = 0 Then Call DuracionPociones(i)
            If .flags.Paralizado = 1 Then Call EfectoParalisisUser(i)
            If .flags.Inmovilizado = 1 Then Call EfectoInmoUser(i)
            If .flags.Ceguera = 1 Then Call EfectoCegueEstu(i)
            If .flags.Estupidez = 1 Then Call EfectoEstupidez(i)
            If .flags.Maldicion = 1 Then Call EfectoMaldicionUser(i)
            If .flags.VelocidadBackup > 0 Then Call EfectoVelocidadUser(i)
        
        End With
        
        If UserList(i).flags.Portal > 1 Then
            UserList(i).flags.Portal = UserList(i).flags.Portal - 1
        
            If UserList(i).flags.Portal = 1 Then
                Mapa = UserList(i).flags.PortalM
                X = UserList(i).flags.PortalX
                Y = UserList(i).flags.PortalY
                Call SendData(SendTarget.toMap, UserList(i).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.TpVerde, 0))
                Call SendData(SendTarget.toMap, UserList(i).flags.PortalM, PrepareMessageLightFXToFloor(X, Y, 0, 105))

                If MapData(Mapa, X, Y).TileExit.Map > 0 Then
                    MapData(Mapa, X, Y).TileExit.Map = 0
                    MapData(Mapa, X, Y).TileExit.X = 0
                    MapData(Mapa, X, Y).TileExit.Y = 0

                End If

                MapData(Mapa, X, Y).Particula = 0
                MapData(Mapa, X, Y).TimeParticula = 0
                MapData(Mapa, X, Y).Particula = 0
                MapData(Mapa, X, Y).TimeParticula = 0
                UserList(i).flags.Portal = 0
                UserList(i).flags.PortalM = 0
                UserList(i).flags.PortalY = 0
                UserList(i).flags.PortalX = 0
                UserList(i).flags.PortalMDestino = 0
                UserList(i).flags.PortalYDestino = 0
                UserList(i).flags.PortalXDestino = 0

            End If

        End If
        
        If UserList(i).Counters.TiempoDeMapeo > 0 Then
            UserList(i).Counters.TiempoDeMapeo = UserList(i).Counters.TiempoDeMapeo - 1

        End If
        
        If UserList(i).flags.Subastando Then
            UserList(i).Counters.TiempoParaSubastar = UserList(i).Counters.TiempoParaSubastar - 1

            If UserList(i).Counters.TiempoParaSubastar = 0 Then
                Call CancelarSubasta

            End If

        End If

        If UserList(i).flags.UserLogged Then

            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                '  If UserList(i).flags.Muerto = 1 Then UserList(i).Counters.Salir = 0
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                ' Call WriteConsoleMsg(i, "Se saldrá del juego en " & UserList(i).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(i, "203", FontTypeNames.FONTTYPE_INFO, UserList(i).Counters.Salir)

                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteDisconnect(i)
                    
                    Call CloseSocket(i)

                End If

            End If

        End If

    Next i

    Exit Sub

errHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)

    Resume Next

End Sub
 
Public Function ReiniciarAutoUpdate() As Double
        
        On Error GoTo ReiniciarAutoUpdate_Err
        

100     ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

        
        Exit Function

ReiniciarAutoUpdate_Err:
        Call RegistrarError(Err.Number, Err.description, "General.ReiniciarAutoUpdate", Erl)
        Resume Next
        
End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
        'WorldSave
        
        On Error GoTo ReiniciarServidor_Err
        
100     Call DoBackUp

        'Guardar Pjs
102     Call GuardarUsuarios
    
104     If EjecutarLauncher Then Shell App.Path & "\launcher.exe" & " megustalanoche*"

        'Chauuu
106     Unload frmMain

        
        Exit Sub

ReiniciarServidor_Err:
        Call RegistrarError(Err.Number, Err.description, "General.ReiniciarServidor", Erl)
        Resume Next
        
End Sub
 
Sub GuardarUsuarios()
        
        On Error GoTo GuardarUsuarios_Err
        
100     haciendoBK = True
    
102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
104     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
        Dim i As Long

106     For i = 1 To LastUser

108         If UserList(i).flags.UserLogged Then
110             If UserList(i).flags.BattleModo = 0 Then
112                 Call SaveUser(i)

                End If

            End If

114     Next i
    
116     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
118     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

120     haciendoBK = False

        
        Exit Sub

GuardarUsuarios_Err:
        Call RegistrarError(Err.Number, Err.description, "General.GuardarUsuarios", Erl)
        Resume Next
        
End Sub

Sub InicializaEstadisticas()
        
        On Error GoTo InicializaEstadisticas_Err
        

        Dim Ta As Long

100     Ta = GetTickCount() And &H7FFFFFFF

102     Call EstadisticasWeb.Inicializa(frmMain.hWnd)
104     Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
106     Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
108     Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
110     Call EstadisticasWeb.Informar(RECORD_USUARIOS, RecordUsuarios)

        
        Exit Sub

InicializaEstadisticas_Err:
        Call RegistrarError(Err.Number, Err.description, "General.InicializaEstadisticas", Erl)
        Resume Next
        
End Sub

Public Sub FreeNPCs()
        
        On Error GoTo FreeNPCs_Err
        

        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all NPC Indexes
        '***************************************************
        Dim LoopC As Long
    
        ' Free all NPC indexes
100     For LoopC = 1 To MAXNPCS
102         Npclist(LoopC).flags.NPCActive = False
104     Next LoopC

        
        Exit Sub

FreeNPCs_Err:
        Call RegistrarError(Err.Number, Err.description, "General.FreeNPCs", Erl)
        Resume Next
        
End Sub

Public Sub FreeCharIndexes()
        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all char indexes
        '***************************************************
        ' Free all char indexes (set them all to 0)
        
        On Error GoTo FreeCharIndexes_Err
        
100     Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

        
        Exit Sub

FreeCharIndexes_Err:
        Call RegistrarError(Err.Number, Err.description, "General.FreeCharIndexes", Erl)
        Resume Next
        
End Sub

Function RandomString(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
        
        On Error GoTo RandomString_Err
        

100     Randomize Time

        Dim rgch As String

102     rgch = "abcdefghijklmnopqrstuvwxyz"
    
104     If OnlyUpper Then
106         rgch = UCase(rgch)
        Else
108         rgch = rgch & UCase(rgch)

        End If
    
110     rgch = rgch & "0123456789"  ' & "#@!~$()-_"

        Dim i As Long

112     For i = 1 To cb
114         RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Next

        
        Exit Function

RandomString_Err:
        Call RegistrarError(Err.Number, Err.description, "General.RandomString", Erl)
        Resume Next
        
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

    On Error GoTo errHnd

    Dim lPos As Long

    Dim lX   As Long

    Dim iAsc As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")

    If (lPos <> 0) Then

        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1

            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))

                If Not CMSValidateChar_(iAsc) Then Exit Function

            End If

        Next lX
        
        'Finale
        CheckMailString = True

    End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
        
        On Error GoTo CMSValidateChar__Err
        
100     CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

        
        Exit Function

CMSValidateChar__Err:
        Call RegistrarError(Err.Number, Err.description, "General.CMSValidateChar_", Erl)
        Resume Next
        
End Function

Public Function Tilde(ByRef data As String) As String

    Dim temp As String

    'Pato
    temp = UCase$(data)
 
    If InStr(1, temp, "Ã") Then temp = Replace$(temp, "Ã", "A")
   
    If InStr(1, temp, "e") Then temp = Replace$(temp, "e", "E")
   
    If InStr(1, temp, "Ã") Then temp = Replace$(temp, "Ã", "I")
   
    If InStr(1, temp, "Ã“") Then temp = Replace$(temp, "Ã“", "O")
   
    If InStr(1, temp, "U") Then temp = Replace$(temp, "U", "U")
   
    Tilde = temp
        
End Function
