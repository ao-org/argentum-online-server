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

Sub DarCuerpoDesnudo(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/14/07
    'Da cuerpo desnudo a un usuario
    '***************************************************
    Dim CuerpoDesnudo As Integer

    Select Case UserList(UserIndex).genero

        Case eGenero.Hombre

            Select Case UserList(UserIndex).raza

                Case eRaza.Humano
                    CuerpoDesnudo = 21 'ok

                Case eRaza.Drow
                    CuerpoDesnudo = 32 ' ok

                Case eRaza.Elfo
                    CuerpoDesnudo = 510 'Revisar

                Case eRaza.Gnomo
                    CuerpoDesnudo = 508 'Revisar

                Case eRaza.Enano
                    CuerpoDesnudo = 53 'ok

                Case eRaza.Orco
                    CuerpoDesnudo = 248 ' ok

            End Select

        Case eGenero.Mujer

            Select Case UserList(UserIndex).raza

                Case eRaza.Humano
                    CuerpoDesnudo = 39 'ok

                Case eRaza.Drow
                    CuerpoDesnudo = 40 'ok

                Case eRaza.Elfo
                    CuerpoDesnudo = 511 'Revisar

                Case eRaza.Gnomo
                    CuerpoDesnudo = 509 'Revisar

                Case eRaza.Enano
                    CuerpoDesnudo = 60 ' ok

                Case eRaza.Orco
                    CuerpoDesnudo = 249 'ok

            End Select

    End Select

    UserList(UserIndex).Char.Body = CuerpoDesnudo

    UserList(UserIndex).flags.Desnudo = 1

End Sub

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal Y As Integer, ByVal b As Boolean)
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(x, Y, b))
    Else
        Call WriteBlockPosition(sndIndex, x, Y, b)

    End If

End Sub

Function HayCosta(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

    'Ladder 10 - 2 - 2010
    'Chequea si hay costa en los tiles proximos al usuario
    If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
        If ((MapData(Map, x, Y).Graphic(1) >= 22552 And MapData(Map, x, Y).Graphic(1) <= 22599) Or (MapData(Map, x, Y).Graphic(1) >= 7283 And MapData(Map, x, Y).Graphic(1) <= 7378) Or (MapData(Map, x, Y).Graphic(1) >= 13387 And MapData(Map, x, Y).Graphic(1) <= 13482)) And MapData(Map, x, Y).Graphic(2) = 0 Then
            HayCosta = True
        Else
            HayCosta = False

        End If

    Else
        HayCosta = False

    End If

End Function

Function HayAgua(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

    If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
        If ((MapData(Map, x, Y).Graphic(1) >= 1505 And MapData(Map, x, Y).Graphic(1) <= 1520) Or (MapData(Map, x, Y).Graphic(1) >= 24223 And MapData(Map, x, Y).Graphic(1) <= 24238) Or (MapData(Map, x, Y).Graphic(1) >= 24303 And MapData(Map, x, Y).Graphic(1) <= 24318) Or (MapData(Map, x, Y).Graphic(1) >= 468 And MapData(Map, x, Y).Graphic(1) <= 483) Or (MapData(Map, x, Y).Graphic(1) >= 44668 And MapData(Map, x, Y).Graphic(1) <= 44939) Or (MapData(Map, x, Y).Graphic(1) >= 24143 And MapData(Map, x, Y).Graphic(1) <= 24158)) And MapData(Map, x, Y).Graphic(2) = 0 Then
            HayAgua = True
        Else
            HayAgua = False

        End If

    Else
        HayAgua = False

    End If

End Function

Private Function HayLava(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer) As Boolean

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/12/07
    '***************************************************
    If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And Y > 0 And Y < 101 Then
        If MapData(Map, x, Y).Graphic(1) >= 5837 And MapData(Map, x, Y).Graphic(1) <= 5852 Then
            HayLava = True
        Else
            HayLava = False

        End If

    Else
        HayLava = False

    End If

End Function

Sub LimpiarMundo()

    '***************************************************
    'Author: Unknow
    'Last Modification: 04/15/2008
    '01/14/2008: Marcos Martinez (ByVal) - La funcion FOR estaba mal. En ves de i habia un 1.
    '04/15/2008: (NicoNZ) - La funcion FOR estaba mal, de la forma que se hacia tiraba error.
    '***************************************************
    On Error GoTo Errhandler

    Dim i As Long

    Dim d As New cGarbage

    For i = TrashCollector.Count To 1 Step -1
        Set d = TrashCollector(i)
        Call EraseObj(1, d.Map, d.x, d.Y)
        Call TrashCollector.Remove(i)
        Set d = Nothing
    Next i

    Call SecurityIp.IpSecurityMantenimientoLista

    Exit Sub

Errhandler:
    Call LogError("Error producido en el sub LimpiarMundo: " & Err.description)

End Sub

Sub LimpiarMundoEntero()

    'Ladder /limpiarmundo
    On Error GoTo Errhandler

    Call GuardarUsuarios

    If BusquedaTesoroActiva Then Exit Sub
    If BusquedaRegaloActiva Then Exit Sub
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpiando mundo....¡Quietos!", FontTypeNames.FONTTYPE_SERVER))

    Dim MapaActual As Long

    Dim Y          As Long

    Dim x          As Long

    For MapaActual = 1 To NumMaps
        For Y = 10 To 91
            For x = 12 To 88

                If MapData(MapaActual, x, Y).ObjInfo.ObjIndex > 0 And MapData(MapaActual, x, Y).Blocked = 0 Then

                    ' If MapData(MapaActual, X, Y).ObjInfo.ObjIndex = 315 Then
                    ' MapData(MapaActual, X, Y).Particula = 0
                    ' MapData(MapaActual, X, Y).TimeParticula = 0
                    'End If
                    If ObjData(MapData(MapaActual, x, Y).ObjInfo.ObjIndex).NoSeLimpia = 0 Then
                        If ItemNoEsDeMapa(MapData(MapaActual, x, Y).ObjInfo.ObjIndex) Then Call EraseObj(10000, MapaActual, x, Y)

                    End If

                End If

            Next x
        Next Y
    Next MapaActual

    LimpiezaTimerMinutos = TimerCleanWorld

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo realizada.", FontTypeNames.FONTTYPE_SERVER))
    Exit Sub
Errhandler:
    Call LogError("Error producido al limpiar las coordenadas " & x & "-" & Y & " del mapa: " & MapaActual & "-" & Err.description)

End Sub

Sub ApagarFogatas()

    'Ladder /ApagarFogatas
    On Error GoTo Errhandler

    Dim obj As obj

    obj.ObjIndex = FOGATA_APAG
    obj.Amount = 1

    Dim MapaActual As Long

    Dim Y          As Long

    Dim x          As Long

    For MapaActual = 1 To NumMaps
        For Y = YMinMapSize To YMaxMapSize
            For x = XMinMapSize To XMaxMapSize

                If MapInfo(MapaActual).lluvia Then
                    If MapData(MapaActual, x, Y).ObjInfo.ObjIndex = FOGATA Then
                        Call EraseObj(10000, MapaActual, x, Y)
                        Call MakeObj(obj, MapaActual, x, Y)

                    End If

                End If

            Next x
        Next Y
    Next MapaActual

    Exit Sub
Errhandler:
    Call LogError("Error producido al apagar las fogatas de " & x & "-" & Y & " del mapa: " & MapaActual & "    -" & Err.description)

End Sub

Sub EnviarSpawnList(ByVal UserIndex As Integer)

    Dim K          As Long

    Dim npcNames() As String

    Debug.Print UBound(SpawnList)
    ReDim npcNames(1 To UBound(SpawnList)) As String

    For K = 1 To UBound(SpawnList)
        npcNames(K) = SpawnList(K).NpcName
    
    Next K

    Call WriteSpawnList(UserIndex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef obj As Object, ByVal Port As Integer)
    #If UsarQueSocket = 0 Then

        obj.AddressFamily = AF_INET
        obj.Protocol = IPPROTO_IP
        obj.SocketType = SOCK_STREAM
        obj.Binary = False
        obj.Blocking = False
        obj.BufferSize = 1024
        obj.LocalPort = Port
        obj.backlog = 5
        obj.listen

    #End If

End Sub

Public Sub LeerLineaComandos()

    Dim rdata As String

    rdata = Command
    rdata = Right$(rdata, Len(rdata))
    ClaveApertura = ReadField(1, rdata, Asc("*")) ' NICK

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
    
    Prision.x = 72
    Prision.Y = 52
    Libertad.x = 73
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
    
    frmCargando.Label1(2).Caption = "Cargando Logros.ini"
    Call CargarLogros ' Ladder 22/04/2015
    
    frmCargando.Label1(2).Caption = "Cargando Baneos Temporales"
    Call LoadBans
    
    frmCargando.Label1(2).Caption = "Cargando Usuarios Donadores"
    Call LoadDonadores
    Call LoadObjDonador
    Call LoadQuests
    
    EstadoGlobal = True
    
    Set Limpieza = New TLimpiezaItem

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
    
    'Conecto base de datos
    Call Database_Connect
    
    'Reinicio los users online
    Call SetUsersLoggedDatabase(0)
    'Tarea pesada
    Call LogoutAllUsersAndAccounts
    
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
    FileExist = LenB(dir$(File, FileType)) <> 0

End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String

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
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)

    End If

End Function

Function MapaValido(ByVal Map As Integer) As Boolean
    MapaValido = Map >= 1 And Map <= NumMaps

End Function

Sub MostrarNumUsers()

    Call SendData(SendTarget.ToAll, 0, PrepareMessageOnlineUser(NumUsers))
    frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
    
    Call SetUsersLoggedDatabase(NumUsers)

End Sub

Public Sub LogCriticEvent(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
    Print #nfile, Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogError(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogConsulta(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\ConsultasGM.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogStatic(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogTarea(Desc As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & Desc
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogClanes(ByVal str As String)

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogDesarrollo(ByVal str As String)

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & str
    Close #nfile

End Sub

Public Sub LogGM(nombre As String, texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
    Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogDatabaseError(Desc As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\Database.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " - " & Desc
    Close #nfile
    
    Exit Sub
    
    Debug.Print Desc
    
Errhandler:

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
    Exit Sub

Errhandler:

End Sub

Public Sub LogAsesinato(texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal

    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub logVentaCasa(ByVal texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal

    Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogHackAttemp(texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogCheating(texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)

    On Error GoTo Errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & Time & " " & texto
    Print #nfile, ""
    Close #nfile

    Exit Sub

Errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean

    Dim Arg As String

    Dim i   As Integer

    For i = 1 To 33

        Arg = ReadField(i, cad, 44)

        If LenB(Arg) = 0 Then Exit Function

    Next i

    ValidInputNP = True

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

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    
    If MapInfo(UserList(UserIndex).Pos.Map).zone <> "DUNGEON" Then
        If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger <> 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger <> 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger < 10 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False

    End If
    
End Function

Public Sub EfectoLluvia(ByVal UserIndex As Integer)

    On Error GoTo Errhandler

    If UserList(UserIndex).flags.UserLogged Then
        If Intemperie(UserIndex) Then

            Dim modifi As Long

            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 3)
            Call QuitarSta(UserIndex, modifi)
            

        End If

    End If

    Exit Sub
Errhandler:
    LogError ("Error en EfectoLluvia")

End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
    
    Dim modifi As Integer
    
    If UserList(UserIndex).Counters.Frio < IntervaloFrio Then
        UserList(UserIndex).Counters.Frio = UserList(UserIndex).Counters.Frio + 1
    Else

        If MapInfo(UserList(UserIndex).Pos.Map).terrain = Nieve Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxHp, 5)
            UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - modifi
            
            If UserList(UserIndex).Stats.MinHp < 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHp = 0
                Call UserDie(UserIndex)

            End If
            
            Call WriteUpdateHP(UserIndex)
        Else
            modifi = Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)
            Call QuitarSta(UserIndex, modifi)

            '  Call WriteUpdateSta(UserIndex)
        End If
        
        UserList(UserIndex).Counters.Frio = 0

    End If

End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/12/07
    'If user is standing on lava, take health points from him
    '***************************************************
    If UserList(UserIndex).Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
        UserList(UserIndex).Counters.Lava = UserList(UserIndex).Counters.Lava + 1
    Else

        If HayLava(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y) Then
            Call WriteConsoleMsg(UserIndex, "¡¡Quitate de la lava, te estás quemando!!.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Porcentaje(UserList(UserIndex).Stats.MaxHp, 5)
            
            If UserList(UserIndex).Stats.MinHp < 1 Then
                Call WriteConsoleMsg(UserIndex, "¡¡Has muerto quemado!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinHp = 0
                Call UserDie(UserIndex)

            End If
            
            Call WriteUpdateHP(UserIndex)

        End If
        
        UserList(UserIndex).Counters.Lava = 0

    End If

End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Invisibilidad > 0 Then
        UserList(UserIndex).Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad - 1
    Else
        UserList(UserIndex).Counters.Invisibilidad = 0
        UserList(UserIndex).flags.invisible = 0

        If UserList(UserIndex).flags.Oculto = 0 Then
            ' Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
            Call WriteContadores(UserIndex)

        End If

    End If

End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

    If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
        Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
    Else
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).flags.Inmovilizado = 0

    End If

End Sub

Public Sub EfectoCegueEstu(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Ceguera > 0 Then
        UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
    Else

        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)

        End If

    End If

End Sub

Public Sub EfectoEstupidez(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Estupidez > 0 Then
        UserList(UserIndex).Counters.Estupidez = UserList(UserIndex).Counters.Estupidez - 1

    Else

        If UserList(UserIndex).flags.Estupidez = 1 Then
            UserList(UserIndex).flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)

        End If

    End If

End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Paralisis > 0 Then
        UserList(UserIndex).Counters.Paralisis = UserList(UserIndex).Counters.Paralisis - 1
    Else
        UserList(UserIndex).flags.Paralizado = 0
        'UserList(UserIndex).Flags.AdministrativeParalisis = 0
        Call WriteParalizeOK(UserIndex)

    End If

End Sub

Public Sub EfectoVelocidadUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Velocidad > 0 Then
        UserList(UserIndex).Counters.Velocidad = UserList(UserIndex).Counters.Velocidad - 1
    Else
        UserList(UserIndex).Char.speeding = UserList(UserIndex).flags.VelocidadBackup
    
        'Call WriteVelocidadToggle(UserIndex)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(UserList(UserIndex).Char.CharIndex, UserList(UserIndex).flags.VelocidadBackup))
        UserList(UserIndex).flags.VelocidadBackup = 0

    End If

End Sub

Public Sub EfectoMaldicionUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Maldicion > 0 Then
        UserList(UserIndex).Counters.Maldicion = UserList(UserIndex).Counters.Maldicion - 1
    
    Else
        UserList(UserIndex).flags.Maldicion = 0
        Call WriteConsoleMsg(UserIndex, "¡La magia perdió su efecto! Ya podes atacar.", FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)

        'Call WriteParalizeOK(UserIndex)
    End If

End Sub

Public Sub EfectoInmoUser(ByVal UserIndex As Integer)

    If UserList(UserIndex).Counters.Inmovilizado > 0 Then
        UserList(UserIndex).Counters.Inmovilizado = UserList(UserIndex).Counters.Inmovilizado - 1
    Else
        UserList(UserIndex).flags.Inmovilizado = 0
        'UserList(UserIndex).Flags.AdministrativeParalisis = 0
        Call WriteInmovilizaOK(UserIndex)

    End If

End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

    Dim massta As Integer

    If UserList(UserIndex).Stats.MinSta < UserList(UserIndex).Stats.MaxSta Then

        If UserList(UserIndex).Counters.STACounter < Intervalo Then
            UserList(UserIndex).Counters.STACounter = UserList(UserIndex).Counters.STACounter + 1
        Else
        
            UserList(UserIndex).Counters.STACounter = 0

            If UserList(UserIndex).flags.Desnudo And Not UserList(UserIndex).flags.Montado Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
            If UserList(UserIndex).Counters.Trabajando > 0 Then Exit Sub  'Trabajando no sube energía. (ToxicWaste)
         
            ' If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub 'Ladder, se ve que esta linea la agregue yo, pero no sirve.

            EnviarStats = True
        
            Dim Suerte As Integer

            If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= -1 Then
                Suerte = 5
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 11 Then
                Suerte = 7
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 21 Then
                Suerte = 9
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 31 Then
                Suerte = 11
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 50 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 41 Then
                Suerte = 13
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 60 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 51 Then
                Suerte = 15
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 70 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 61 Then
                Suerte = 17
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 80 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 71 Then
                Suerte = 19
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 90 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 81 Then
                Suerte = 21
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 100 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 91 Then
                Suerte = 23
            ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) = 100 Then
                Suerte = 25

            End If
        
            If UserList(UserIndex).flags.RegeneracionSta = 1 Then
                Suerte = 45

            End If
        
            massta = RandomNumber(1, Porcentaje(UserList(UserIndex).Stats.MaxSta, Suerte))
            UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta + massta

            If UserList(UserIndex).Stats.MinSta > UserList(UserIndex).Stats.MaxSta Then
                UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MaxSta

            End If

        End If

    End If

End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)

    Dim n As Integer

    If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
        UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
    Else
        'Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas moriras.", FontTypeNames.FONTTYPE_VENENO)
        Call WriteLocaleMsg(UserIndex, "47", FontTypeNames.FONTTYPE_VENENO)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Envenena, 30, False))
        UserList(UserIndex).Counters.Veneno = 0
        n = RandomNumber(3, 6)
        n = n * UserList(UserIndex).flags.Envenenado
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - n

        If UserList(UserIndex).Stats.MinHp < 1 Then Call UserDie(UserIndex)
        Call WriteUpdateHP(UserIndex)

    End If

End Sub

Public Sub EfectoAhogo(ByVal UserIndex As Integer)

    Dim n As Integer

    If RequiereOxigeno(UserList(UserIndex).Pos.Map) Then
        If UserList(UserIndex).Counters.Ahogo < 70 Then
            UserList(UserIndex).Counters.Ahogo = UserList(UserIndex).Counters.Ahogo + 1
        Else
            Call WriteConsoleMsg(UserIndex, "Te estas ahogando.. si no consigues oxigeno moriras.", FontTypeNames.FONTTYPE_EJECUCION)
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 205, 30, False))
            UserList(UserIndex).Counters.Ahogo = 0
            n = RandomNumber(150, 200)
            UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - n

            If UserList(UserIndex).Stats.MinHp < 1 Then
                Call UserDie(UserIndex)
                UserList(UserIndex).flags.Ahogandose = 0

            End If

            Call WriteUpdateHP(UserIndex)

        End If

    Else
        UserList(UserIndex).flags.Ahogandose = 0

    End If

End Sub

Public Sub EfectoIncineramiento(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean)

    Dim n As Integer
 
    If UserList(UserIndex).Counters.Incineracion < IntervaloIncineracion Then
        UserList(UserIndex).Counters.Incineracion = UserList(UserIndex).Counters.Incineracion + 1
    Else
        Call WriteConsoleMsg(UserIndex, "Te estas incinerando,si no te curas moriras.", FontTypeNames.FONTTYPE_INFO)
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Incinerar, 30, False))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 73, 0))
        UserList(UserIndex).Counters.Incineracion = 0
        n = RandomNumber(40, 80)
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - n

        If UserList(UserIndex).Stats.MinHp < 1 Then Call UserDie(UserIndex)
        Call WriteUpdateHP(UserIndex)

    End If
 
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)

    'Controla la duracion de las pociones
    If UserList(UserIndex).flags.DuracionEfecto > 0 Then
        UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1

        If UserList(UserIndex).flags.DuracionEfecto = 0 Then
            UserList(UserIndex).flags.TomoPocion = False
            UserList(UserIndex).flags.TipoPocion = 0

            'volvemos los atributos al estado normal
            Dim LoopX As Integer

            For LoopX = 1 To NUMATRIBUTOS
                UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
            Next
            Call WriteFYA(UserIndex)

        End If

    End If

End Sub

Public Sub HambreYSed(ByVal UserIndex As Integer, ByRef fenviarAyS As Boolean)

    If Not UserList(UserIndex).flags.Privilegios And PlayerType.user Then Exit Sub
    If UserList(UserIndex).flags.BattleModo = 1 Then Exit Sub

    'Sed
    If UserList(UserIndex).Stats.MinAGU > 0 Then
        If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
            UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
        Else
            UserList(UserIndex).Counters.AGUACounter = 0
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
            If UserList(UserIndex).Stats.MinAGU <= 0 Then
                UserList(UserIndex).Stats.MinAGU = 0
                UserList(UserIndex).flags.Sed = 1

            End If
        
            fenviarAyS = True

        End If

    End If

    'hambre
    If UserList(UserIndex).Stats.MinHam > 0 Then
        If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
            UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
        Else
            UserList(UserIndex).Counters.COMCounter = 0
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10

            If UserList(UserIndex).Stats.MinHam <= 0 Then
                UserList(UserIndex).Stats.MinHam = 0
                UserList(UserIndex).flags.Hambre = 1

            End If

            fenviarAyS = True

        End If

    End If

End Sub

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

    If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

    Dim mashit As Integer

    'con el paso del tiempo va sanando....pero muy lentamente ;-)
    If UserList(UserIndex).Stats.MinHp < UserList(UserIndex).Stats.MaxHp Then
        If UserList(UserIndex).flags.RegeneracionHP = 1 Then
            Intervalo = 400

        End If
    
        If UserList(UserIndex).Counters.HPCounter < Intervalo Then
            UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
        Else
            mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5))
        
            UserList(UserIndex).Counters.HPCounter = 0
            UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + mashit

            If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
            Call WriteConsoleMsg(UserIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
            EnviarStats = True

        End If

    End If

End Sub

Public Sub CargaNpcsDat()

    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
    
    'npcfile = DatPath & "NPCs-HOSTILES.dat"
    'Call LeerNPCsHostiles.Initialize(npcfile)
End Sub

Sub PasarSegundo()

    On Error GoTo Errhandler

    Dim i    As Long

    Dim h    As Byte

    Dim Mapa As Integer

    Dim x    As Byte

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
                x = UserList(i).flags.PortalX
                Y = UserList(i).flags.PortalY
                Call SendData(SendTarget.toMap, UserList(i).flags.PortalM, PrepareMessageParticleFXToFloor(x, Y, ParticulasIndex.TpVerde, 0))
                Call SendData(SendTarget.toMap, UserList(i).flags.PortalM, PrepareMessageLightFXToFloor(x, Y, 0, 105))

                If MapData(Mapa, x, Y).TileExit.Map > 0 Then
                    MapData(Mapa, x, Y).TileExit.Map = 0
                    MapData(Mapa, x, Y).TileExit.x = 0
                    MapData(Mapa, x, Y).TileExit.Y = 0

                End If

                MapData(Mapa, x, Y).Particula = 0
                MapData(Mapa, x, Y).TimeParticula = 0
                MapData(Mapa, x, Y).Particula = 0
                MapData(Mapa, x, Y).TimeParticula = 0
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

Errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)

    Resume Next

End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell App.Path & "\launcher.exe" & " megustalanoche*"

    'Chauuu
    Unload frmMain

End Sub
 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Long

    For i = 1 To LastUser

        If UserList(i).flags.UserLogged Then
            If UserList(i).flags.BattleModo = 0 Then
                Call SaveUser(i)

            End If

        End If

    Next i
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False

End Sub

Sub InicializaEstadisticas()

    Dim Ta As Long

    Ta = GetTickCount() And &H7FFFFFFF

    Call EstadisticasWeb.Inicializa(frmMain.hWnd)
    Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all NPC Indexes
    '***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC

End Sub

Public Sub FreeCharIndexes()
    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Releases all char indexes
    '***************************************************
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

End Sub

Function RandomString(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String

    Randomize Time

    Dim rgch As String

    rgch = "abcdefghijklmnopqrstuvwxyz"
    
    If OnlyUpper Then
        rgch = UCase(rgch)
    Else
        rgch = rgch & UCase(rgch)

    End If
    
    rgch = rgch & "0123456789"  ' & "#@!~$()-_"

    Dim i As Long

    For i = 1 To cb
        RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

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
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

End Function
