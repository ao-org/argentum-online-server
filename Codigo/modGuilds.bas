Attribute VB_Name = "modGuilds"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE             As String
'archivo .\guilds\guildinfo.ini o similar
Private Const MAX_GUILDS          As Integer = 1000
'cantidad maxima de guilds en el servidor
Public CANTIDADDECLANES           As Integer
'cantidad actual de clanes en el servidor
Private guilds(1 To MAX_GUILDS)   As clsClan
'array global de guilds, se indexa por userlist().guildindex
Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir
Public Const MAXASPIRANTES        As Byte = 10

'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion
Public Enum e_ALINEACION_GUILD
    ALINEACION_NEUTRAL = 0
    ALINEACION_ARMADA = 1
    ALINEACION_CAOTICA = 2
    ALINEACION_CIUDADANA = 3
    ALINEACION_CRIMINAL = 4
End Enum

'numero de .wav del cliente
Public Enum e_RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum

'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadGuildsDB()
    On Error GoTo LoadGuildsDB_Err
    Dim CantClanes As String
    Dim i          As Integer
    Dim TempStr    As String
    Dim Alin       As e_ALINEACION_GUILD
    Dim RS         As Recordset
    Set RS = Query("SELECT id, founder_id, guild_name, creation_date, alignment, last_elections, description, news, leader_id, level, current_exp, flag_file FROM guilds")
    If RS Is Nothing Then Exit Sub
    CANTIDADDECLANES = RS.RecordCount
    i = 0
    If Not RS.RecordCount = 0 Then
        While Not RS.EOF
            i = i + 1
            Set guilds(i) = New clsClan
            Call guilds(i).InitFromRecord(RS, i)
            RS.MoveNext
        Wend
    End If
    Exit Sub
LoadGuildsDB_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.LoadGuildsDB", Erl)
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo m_ConectarMiembroAClan_Err
    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    End If
    Exit Function
m_ConectarMiembroAClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_ConectarMiembroAClan", Erl)
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    On Error GoTo m_DesconectarMiembroDelClan_Err
    If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(UserIndex)
    Exit Sub
m_DesconectarMiembroDelClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_DesconectarMiembroDelClan", Erl)
End Sub

Private Function m_EsGuildLeader(ByRef UserId As Long, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo m_EsGuildLeader_Err
    m_EsGuildLeader = (UserId = guilds(GuildIndex).GetLeader)
    Exit Function
m_EsGuildLeader_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_EsGuildLeader", Erl)
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal ExpellUserId As Long) As Integer
    On Error GoTo m_EcharMiembroDeClan_Err
    'UI echa a Expulsado del clan de Expulsado
    Dim UserReference As t_UserReference
    Dim GI            As Integer
    Dim Map           As Integer
    Dim ExpelledName  As String
    m_EcharMiembroDeClan = 0
    ExpelledName = GetUserName(ExpellUserId)
    UserReference = NameIndex(ExpelledName)
    If IsValidUserRef(UserReference) Then
        'pj online
        GI = UserList(UserReference.ArrayIndex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(ExpellUserId, GI, Expulsador) Then
                If m_EsGuildLeader(ExpellUserId, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).DesConectarMiembro(UserReference.ArrayIndex)
                Call guilds(GI).ExpulsarMiembro(ExpellUserId)
                Call LogClanes(ExpelledName & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(UserReference.ArrayIndex).GuildIndex = 0
                Map = UserList(UserReference.ArrayIndex).pos.Map
                If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
                    Call WarpUserChar(UserReference.ArrayIndex, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.x, MapInfo(Map).Salida.y, True)
                    Call WriteConsoleMsg(UserReference.ArrayIndex, PrepareMessageLocaleMsg(1941, vbNullString, e_FontTypeNames.FONTTYPE_GUILDMSG)) ' Msg1941=Necesitas un clan para pertenecer en este mapa.
                Else
                    Call RefreshCharStatus(UserReference.ArrayIndex)
                End If
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetUserGuildIndexDatabase(ExpellUserId)
        If GI > 0 Then
            If m_PuedeSalirDeClan(ExpellUserId, GI, Expulsador) Then
                If m_EsGuildLeader(ExpellUserId, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).ExpulsarMiembro(ExpellUserId)
                Call LogClanes(ExpelledName & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                Map = GetMapDatabase(ExpelledName)
                If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
                    Call SetPositionDatabase(ExpelledName, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.x, MapInfo(Map).Salida.y)
                End If
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If
    Exit Function
m_EcharMiembroDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_EcharMiembroDeClan", Erl)
End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
    On Error GoTo ActualizarWebSite_Err
    Dim GI As Integer
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    If Not m_EsGuildLeader(UserList(UserIndex).Id, GI) Then Exit Sub
    Call guilds(GI).SetURL(Web)
    Exit Sub
ActualizarWebSite_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.ActualizarWebSite", Erl)
End Sub

Public Sub ChangeCodexAndDesc(ByRef Desc As String, ByVal GuildIndex As Integer)
    On Error GoTo ChangeCodexAndDesc_Err
    Dim i As Long
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    With guilds(GuildIndex)
        Call .SetDesc(Desc)
    End With
    Exit Sub
ChangeCodexAndDesc_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.ChangeCodexAndDesc", Erl)
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
    On Error GoTo ActualizarNoticias_Err
    Dim GI As Integer
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    If Not m_EsGuildLeader(UserList(UserIndex).Id, GI) Then Exit Sub
    Call guilds(GI).SetGuildNews(Datos)
    Exit Sub
ActualizarNoticias_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.ActualizarNoticias", Erl)
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, _
                               ByRef Desc As String, _
                               ByRef GuildName As String, _
                               ByVal Alineacion As e_ALINEACION_GUILD, _
                               ByRef refError As String) As Boolean
    On Error GoTo CrearNuevoClan_Err
    Dim i           As Integer
    Dim DummyString As String
    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If
    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = 2020 'Nombre de clan inválido.
        Exit Function
    End If
    If YaExiste(GuildName) Then
        refError = 2021 'Ya existe un clan con ese nombre.
        Exit Function
    End If
    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan
        'constructor custom de la clase clan
        Set guilds(CANTIDADDECLANES) = New clsClan
        'Damos de alta al clan como nuevo inicializando sus archivos
        Call guilds(CANTIDADDECLANES).InicializarNuevoClan(GuildName, CANTIDADDECLANES, Alineacion, UserList(FundadorIndex).Id)
        'seteamos codex y descripcion
        Call guilds(CANTIDADDECLANES).SetDesc(Desc)
        Call guilds(CANTIDADDECLANES).SetGuildNews("¡Bienvenido a " & GuildName & "! Clan creado con alineación : " & Alineacion2String(Alineacion) & ".")
        Call guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).Id)
        Call guilds(CANTIDADDECLANES).SetNivelDeClan(1)
        Call guilds(CANTIDADDECLANES).SetExpActual(0)
        '"conectamos" al nuevo miembro a la lista de la clase
        Call guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).Id)
        Call guilds(CANTIDADDECLANES).ConectarMiembro(FundadorIndex)
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        For i = 1 To CANTIDADDECLANES - 1
            Call guilds(i).ProcesarFundacionDeOtroClan
        Next i
    Else
        refError = 2022 'No hay más slots para fundar clanes. Consulte a un administrador.
        Exit Function
    End If
    CrearNuevoClan = True
    Exit Function
CrearNuevoClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.CrearNuevoClan", Erl)
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer, ByRef guildList() As String)
    On Error GoTo SendGuildNews_Err
    Dim GuildIndex As Integer
    Dim i          As Integer
    Dim go         As Integer
    Dim ClanNivel  As Byte
    Dim ExpAcu     As Integer
    Dim ExpNe      As Integer
    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex = 0 Then Exit Sub
    Dim MemberList() As Long
    MemberList = guilds(GuildIndex).GetMemberList()
    ClanNivel = guilds(GuildIndex).GetNivelDeClan
    ExpAcu = guilds(GuildIndex).GetExpActual
    ExpNe = GetRequiredExpForGuildLevel(ClanNivel)
    Call WriteGuildNews(UserIndex, guilds(GuildIndex).GetGuildNews, guildList, MemberList, ClanNivel, ExpAcu, ExpNe)
    Exit Sub
SendGuildNews_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SendGuildNews", Erl)
End Sub

Public Function m_PuedeSalirDeClan(ByRef UserId As Long, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
    'sale solo si no es fundador del clan.
    On Error GoTo m_PuedeSalirDeClan_Err
    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Or guilds(GuildIndex) Is Nothing Then Exit Function
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If
    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios And e_PlayerType.User Then
        If Not m_EsGuildLeader(UserList(QuienLoEchaUI).Id, GuildIndex) Then
            If UserList(QuienLoEchaUI).Id <> UserId Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If
    m_PuedeSalirDeClan = guilds(GuildIndex).Fundador <> UserId
    Exit Function
m_PuedeSalirDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_PuedeSalirDeClan", Erl)
End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As e_ALINEACION_GUILD, ByRef refError As String) As Boolean
    On Error GoTo PuedeFundarUnClan_Err
    PuedeFundarUnClan = False
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = 2023 'Ya perteneces a un clan, no podés fundar otro.
        Exit Function
    End If
    If UserList(UserIndex).Stats.ELV < 23 Or UserList(UserIndex).Stats.UserSkills(e_Skill.liderazgo) < 50 Then
        refError = 2024 'Para fundar un clan debes ser Nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar.
        Exit Function
    End If
    If Not TieneObjetos(407, 1, UserIndex) Then
        refError = 2025 'Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar.
        Exit Function
    End If
    If Not TieneObjetos(408, 1, UserIndex) Then
        refError = 2026 'Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar.
        Exit Function
    End If
    If Not TieneObjetos(409, 1, UserIndex) Then
        refError = 2027 'Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar.
        Exit Function
    End If
    If Not TieneObjetos(412, 1, UserIndex) Then
        refError = 2028 'Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar.
        Exit Function
    End If
    If Alineacion = e_ALINEACION_GUILD.ALINEACION_CIUDADANA And UserList(UserIndex).flags.Seguro = False Then
        refError = 2029 'Para fundar un clan ciudadano deberás tener activado el seguro.
        Exit Function
    End If
    Select Case Alineacion
        Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
            If Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Or Status(UserIndex) = _
                    e_Facciones.concilio Then
                refError = 2030 'Para fundar un clan neutral deberás ser ciudadano o criminal.
                Exit Function
            End If
        Case e_ALINEACION_GUILD.ALINEACION_ARMADA
            If Status(UserIndex) <> e_Facciones.Armada And Status(UserIndex) <> e_Facciones.consejo Then
                refError = 2031 'Para fundar un clan de la Armada Real deberás pertenecer a la misma.
                Exit Function
            End If
        Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
            If Status(UserIndex) <> e_Facciones.Caos And Status(UserIndex) <> e_Facciones.concilio Then
                refError = 2032 'Para fundar un clan de la Legión Oscura deberás pertenecer a la misma.
                Exit Function
            End If
        Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
            If Status(UserIndex) <> e_Facciones.Ciudadano And Status(UserIndex) <> e_Facciones.Armada Then
                refError = 2033 'Para fundar un clan ciudadano deberás ser ciudadano.
                Exit Function
            End If
        Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Status(UserIndex) <> e_Facciones.Criminal And Status(UserIndex) <> e_Facciones.Caos Then
                refError = 2034 'Para fundar un clan criminal deberás ser criminal o legión oscura.
                Exit Function
            End If
    End Select
    PuedeFundarUnClan = True
    Exit Function
PuedeFundarUnClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.PuedeFundarUnClan", Erl)
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo m_EstadoPermiteEntrarChar_Err
    Dim Promedio As Long
    Dim ELV      As Integer
    Dim f        As Byte
    m_EstadoPermiteEntrarChar = False
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)
    End If
    If PersonajeExiste(Personaje) Then
        Dim Status As Integer
        Status = CInt(GetUserValue(LCase$(Personaje), "status"))
        Select Case guilds(GuildIndex).Alineacion
            Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
                m_EstadoPermiteEntrarChar = (Status = e_Facciones.Ciudadano Or Status = e_Facciones.Criminal)
            Case e_ALINEACION_GUILD.ALINEACION_ARMADA
                m_EstadoPermiteEntrarChar = (Status = e_Facciones.Armada Or Status = e_Facciones.consejo)
            Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
                m_EstadoPermiteEntrarChar = (Status = e_Facciones.Caos Or Status = e_Facciones.concilio)
            Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
                m_EstadoPermiteEntrarChar = (Status = e_Facciones.Ciudadano Or Status = e_Facciones.Armada)
            Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrarChar = (Status = e_Facciones.Criminal Or Status = e_Facciones.Caos)
        End Select
    End If
    Exit Function
m_EstadoPermiteEntrarChar_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrarChar", Erl)
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo m_EstadoPermiteEntrar_Err
    Select Case guilds(GuildIndex).Alineacion
        Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
            m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Ciudadano Or Status(UserIndex) = e_Facciones.Criminal
        Case e_ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo
        Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
            m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio
        Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
            m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Ciudadano Or Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo
        Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Criminal Or Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio
    End Select
    Exit Function
m_EstadoPermiteEntrar_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrar", Erl)
End Function

Public Function Alineacion2String(ByVal Alineacion As e_ALINEACION_GUILD) As String
    On Error GoTo Alineacion2String_Err
    Select Case Alineacion
        Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
            Alineacion2String = "Neutral"
        Case e_ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Armada Real"
        Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
            Alineacion2String = "Legión Oscura"
        Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
            Alineacion2String = "Ciudadano"
        Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"
    End Select
    Exit Function
Alineacion2String_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.Alineacion2String", Erl)
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
    On Error GoTo GuildNameValido_Err
    Dim car As Byte
    Dim i   As Integer
    'old function by morgo
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            GuildNameValido = False
            Exit Function
        End If
    Next i
    GuildNameValido = True
    Exit Function
GuildNameValido_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GuildNameValido", Erl)
End Function

Public Function YaExiste(ByVal GuildName As String) As Boolean
    On Error GoTo YaExiste_Err
    Dim i As Integer
    YaExiste = False
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
        If YaExiste Then Exit Function
    Next i
    Exit Function
YaExiste_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.YaExiste", Erl)
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
    On Error GoTo GuildIndex_Err
    'me da el indice del guildname
    Dim i As Integer
    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
    Exit Function
GuildIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GuildIndex", Erl)
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
    On Error GoTo m_ListaDeMiembrosOnline_Err
    Dim i As Integer
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And ((UserList(i).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) <> 0 Or (UserList( _
                    UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) <> 0)) Then m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).name & _
                    ","
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
    Exit Function
m_ListaDeMiembrosOnline_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_ListaDeMiembrosOnline", Erl)
End Function

Public Function PrepareGuildsList() As String()
    On Error GoTo PrepareGuildsList_Err
    Dim tStr() As String
    Dim i      As Long
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        For i = 1 To CANTIDADDECLANES
            tStr(i - 1) = guilds(i).GuildName & "-" & guilds(i).Alineacion
        Next i
    End If
    PrepareGuildsList = tStr
    Exit Function
PrepareGuildsList_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.PrepareGuildsList", Erl)
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
    On Error GoTo SendGuildDetails_Err
    Dim codex(CANTIDADMAXIMACODEX - 1) As String
    Dim GI                             As Integer
    Dim i                              As Long
    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Sub
    With guilds(GI)
        Call WriteGuildDetails(UserIndex, GuildName, GetUserName(.Fundador), .GetFechaFundacion, .GetLeader, .CantidadDeMiembros, Alineacion2String(.Alineacion), .GetDesc, _
                .GetNivelDeClan)
    End With
    Exit Sub
SendGuildDetails_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SendGuildDetails", Erl)
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
    On Error GoTo SendGuildLeaderInfo_Err
    Dim GI              As Integer
    Dim guildList()     As String
    Dim MemberList()    As Long
    Dim aspirantsList() As String
    With UserList(UserIndex)
        GI = .GuildIndex
        guildList = PrepareGuildsList()
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            'Send the guild list instead
            Call WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        If Not m_EsGuildLeader(.Id, GI) Then
            'Send the guild list instead
            Call modGuilds.SendGuildNews(UserIndex, guildList)
            Exit Sub
        End If
        MemberList = guilds(GI).GetMemberList()
        aspirantsList = guilds(GI).GetAspirantes()
        Dim ClanLevel As Integer
        ClanLevel = guilds(GI).GetNivelDeClan
        Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList, ClanLevel, guilds(GI).GetExpActual, GetRequiredExpForGuildLevel( _
                ClanLevel))
    End With
    Exit Sub
SendGuildLeaderInfo_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SendGuildLeaderInfo", Erl)
End Sub

Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    On Error GoTo m_Iterador_ProximoUserIndex_Err
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
    Exit Function
m_Iterador_ProximoUserIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.m_Iterador_ProximoUserIndex", Erl)
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    On Error GoTo Iterador_ProximoGM_Err
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
    End If
    Exit Function
Iterador_ProximoGM_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.Iterador_ProximoGM", Erl)
End Function

Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
    On Error GoTo GMEscuchaClan_Err
    Dim GI As Integer
    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteLocaleMsg(UserIndex, 1603, guilds(UserList(UserIndex).EscucheClan).GuildName, e_FontTypeNames.FONTTYPE_GUILD) 'Msg1603= Dejas de escuchar a : ¬1
        guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
        Exit Function
    End If
    'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(UserIndex).EscucheClan <> 0 Then
            If UserList(UserIndex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1942, GuildName, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1942=Conectado a : ¬1
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1943, guilds(UserList(UserIndex).EscucheClan).GuildName, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1943=Dejas de escuchar a : ¬1
                guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            End If
        End If
        Call guilds(GI).ConectarGM(UserIndex)
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1942, GuildName, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1942=Conectado a : ¬1
        GMEscuchaClan = GI
        UserList(UserIndex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1944, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1944=Error, el clan no existe.
        GMEscuchaClan = 0
    End If
    Exit Function
GMEscuchaClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GMEscuchaClan", Erl)
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    'el index lo tengo que tener de cuando me puse a escuchar
    On Error GoTo GMDejaDeEscucharClan_Err
    UserList(UserIndex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(UserIndex)
    Exit Sub
GMDejaDeEscucharClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GMDejaDeEscucharClan", Erl)
End Sub

Public Function PersonajeEsLeader(ByVal CharId As Long) As Boolean
    Dim GuildIndex As Integer
    GuildIndex = GetUserGuildIndexDatabase(CharId)
    If GuildIndex > 0 Then
        If m_EsGuildLeader(CharId, GuildIndex) Then PersonajeEsLeader = True
    End If
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByRef Detalles As String)
    On Error GoTo a_RechazarAspiranteChar_Err
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    Call SaveUserGuildRejectionReason(Aspirante, Detalles)
    Exit Sub
a_RechazarAspiranteChar_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.a_RechazarAspiranteChar", Erl)
End Sub

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef nombre As String, ByRef refError As String) As Boolean
    On Error GoTo a_RechazarAspirante_Err
    Dim GI           As Integer
    Dim NroAspirante As Integer
    a_RechazarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = 2035 'No perteneces a ningún clan.
        Exit Function
    End If
    Call guilds(GI).RetirarAspirante(nombre)
    refError = 2036 'Fue rechazada tu solicitud de ingreso a ¬1.
    a_RechazarAspirante = True
    Exit Function
a_RechazarAspirante_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.a_RechazarAspirante", Erl)
End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef name As String) As String
    On Error GoTo a_DetallesAspirante_Err
    Dim GI           As Integer
    Dim NroAspirante As Integer
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(UserIndex).Id, GI) Then
        Exit Function
    End If
    a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(name)
    Exit Function
a_DetallesAspirante_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.a_DetallesAspirante", Erl)
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
    Dim GI     As Integer
    Dim NroAsp As Integer
    Dim list() As Long
    Dim i      As Long
    On Error GoTo Error
    GI = UserList(UserIndex).GuildIndex
    Personaje = UCase$(Personaje)
    If Not PersonajeExiste(Personaje) Then
        Call guilds(GI).ExpulsarMiembro(Personaje)
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1945, vbNullString, e_FontTypeNames.FONTTYPE_GUILDMSG)) ' Msg1945=El personaje no existe y fue eliminado de la lista de miembros.
        Exit Sub
    End If
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1946, vbNullString, e_FontTypeNames.FONTTYPE_GUILDMSG)) ' Msg1946=No perteneces a ningún clan.
        Exit Sub
    End If
    If Not m_EsGuildLeader(UserList(UserIndex).Id, GI) Then
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1947, vbNullString, e_FontTypeNames.FONTTYPE_GUILDMSG)) ' Msg1947=No eres el líder de tu clan.
        Exit Sub
    End If
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    Dim HasRequest As Boolean
    Dim CharId     As Long
    CharId = GetCharacterIdWithName(Personaje)
    HasRequest = guilds(GI).HasGuildRequest(CharId)
    If Not HasRequest Then
        list = guilds(GI).GetMemberList()
        For i = 0 To UBound(list())
            If CharId = list(i) Then Exit For
        Next i
        If i > UBound(list()) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1948, vbNullString, e_FontTypeNames.FONTTYPE_GUILDMSG)) ' Msg1948=El personaje no es ni aspirante ni miembro del clan.
            Exit Sub
        End If
    End If
    Call SendCharacterInfoDatabase(UserIndex, Personaje)
    Exit Sub
Error:
    If Not PersonajeExiste(Personaje) Then
        Call LogError("El usuario " & UserList(UserIndex).name & " (" & UserIndex & " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
    Else
        Call LogError("[" & Err.Number & "] " & Err.Description & " En la rutina SendDetallesPersonaje, por el usuario " & UserList(UserIndex).name & " (" & UserIndex & _
                " ), pidiendo informacion sobre el personaje " & Personaje)
    End If
End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
    On Error GoTo a_NuevoAspirante_Err
    Dim ViejoSolicitado   As String
    Dim ViejoGuildINdex   As Integer
    Dim ViejoNroAspirante As Integer
    Dim NuevoGuildIndex   As Integer
    a_NuevoAspirante = False
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = 2010 'Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro.
        Exit Function
    End If
    If EsNewbie(UserIndex) Then
        refError = 2005 'Los newbies no tienen derecho a entrar a un clan.
        Exit Function
    End If
    NuevoGuildIndex = GuildIndex(clan)
    If NuevoGuildIndex = 0 Then
        refError = 2006 'Ese clan no existe! Avise a un administrador.
        Exit Function
    End If
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = 2007 & "¬" & Alineacion2String(guilds(NuevoGuildIndex).Alineacion) 'Tú no podés entrar a un clan de alineación ¬1.
        Exit Function
    End If
    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = 2008 'El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes.
        Exit Function
    End If
    Dim NuevoGuildAspirantes() As String
    NuevoGuildAspirantes = guilds(NuevoGuildIndex).GetAspirantes()
    Dim i As Long
    For i = 0 To UBound(NuevoGuildAspirantes)
        If UserList(UserIndex).name = NuevoGuildAspirantes(i) Then
            refError = 2009 'Ya has enviado una solicitud a este clan.
            Exit Function
        End If
    Next
    ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).name & ".chr", "GUILD", "ASPIRANTEA")
    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).name)
        End If
    End If
    Call SendData(SendTarget.ToDiosesYclan, NuevoGuildIndex, PrepareMessageGuildChat("Msg2039¬" & UserList(UserIndex).name, 7))  'Msg2039=Clan: [¬1] ha enviado solicitud para unirse al clan.
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).name, Solicitud)
    a_NuevoAspirante = True
    Exit Function
a_NuevoAspirante_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.a_NuevoAspirante", Erl)
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
    On Error GoTo a_AceptarAspirante_Err
    Dim GI           As Integer
    Dim tGI          As Integer
    Dim AspiranteRef As t_UserReference
    'un pj ingresa al clan :D
    a_AceptarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = 2011 'No perteneces a ningún clan.
        Exit Function
    End If
    If Not m_EsGuildLeader(UserList(UserIndex).Id, GI) Then
        refError = 2012 'No eres el líder de tu clan.
        Exit Function
    End If
    Dim UserDidRequest As Boolean
    Dim CharId         As Long
    CharId = GetCharacterIdWithName(Aspirante)
    UserDidRequest = guilds(GI).HasGuildRequest(CharId)
    If Not UserDidRequest Then
        refError = 2013 'El Pj no es aspirante al clan.
        Exit Function
    End If
    AspiranteRef = NameIndex(Aspirante)
    If IsValidUserRef(AspiranteRef) Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteRef.ArrayIndex, GI) Then
            refError = 2014 & "¬" & Aspirante & "¬" & Alineacion2String(guilds(GI).Alineacion) '¬1 no puede entrar a un clan ¬2.
            Call guilds(GI).RetirarAspirante(Aspirante)
            Exit Function
        ElseIf Not UserList(AspiranteRef.ArrayIndex).GuildIndex = 0 Then
            refError = 2015 & "¬" & Aspirante '¬1 ya es parte de otro clan.
            Call guilds(GI).RetirarAspirante(Aspirante)
            Exit Function
        End If
        If GuildAlignmentIndex(GI) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA And UserList(AspiranteRef.ArrayIndex).flags.Seguro = False Then
            refError = 2016 & "¬" & Aspirante '¬1 deberá activar el seguro para entrar al clan.
            Call guilds(GI).RetirarAspirante(Aspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = 2017 & "¬" & Aspirante & "¬" & Alineacion2String(guilds(GI).Alineacion) '¬1 no puede entrar a un clan ¬2.
            Call guilds(GI).RetirarAspirante(Aspirante)
            Exit Function
        Else
            tGI = GetUserGuildIndexDatabase(CharId)
            If tGI <> 0 Then
                refError = 2018 & "¬" & Aspirante '¬1 ya es parte de otro clan.
                Call guilds(GI).RetirarAspirante(Aspirante)
                Exit Function
            End If
        End If
    End If
    If guilds(GI).CantidadDeMiembros >= MiembrosPermite(GI) Then
        refError = 2019 'La capacidad del clan está completa.
        Exit Function
    End If
    'el pj es aspirante al clan y puede entrar
    Call guilds(GI).RetirarAspirante(Aspirante)
    Call guilds(GI).AceptarNuevoMiembro(CharId)
    ' If player is online, update tag
    If IsValidUserRef(AspiranteRef) Then
        Call RefreshCharStatus(AspiranteRef.ArrayIndex)
    End If
    a_AceptarAspirante = True
    Exit Function
a_AceptarAspirante_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.a_AceptarAspirante", Erl)
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    On Error GoTo GuildName_Err
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildName = guilds(GuildIndex).GuildName
    Exit Function
GuildName_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GuildName", Erl)
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    On Error GoTo GuildLeader_Err
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildLeader = guilds(GuildIndex).GetLeader
    Exit Function
GuildLeader_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GuildLeader", Erl)
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    On Error GoTo GuildAlignment_Err
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
    Exit Function
GuildAlignment_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GuildAlignment", Erl)
End Function

Public Function GuildAlignmentIndex(ByVal GuildIndex As Integer) As e_ALINEACION_GUILD
    On Error GoTo GuildAlignmentIndex_Err
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    GuildAlignmentIndex = guilds(GuildIndex).Alineacion
    Exit Function
GuildAlignmentIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GuildAlignmentIndex", Erl)
End Function

Public Function NivelDeClan(ByVal GuildIndex As Integer) As Byte
    On Error GoTo NivelDeClan_Err
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    NivelDeClan = guilds(GuildIndex).GetNivelDeClan
    Exit Function
NivelDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.NivelDeClan", Erl)
End Function

Public Function Alineacion(ByVal GuildIndex As Integer) As Byte
    On Error GoTo Alineacion_Err
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    Alineacion = guilds(GuildIndex).Alineacion
    Exit Function
Alineacion_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.Alineacion", Erl)
End Function

Sub CheckClanExp(ByVal UserIndex As Integer, ByVal ExpDar As Integer)
    On Error GoTo CheckClanExp_Err
    Dim ExpActual    As Integer
    Dim ExpNecesaria As Integer
    Dim GI           As Integer
    Dim nivel        As Byte
    With UserList(UserIndex)
        GI = .GuildIndex
        ExpActual = guilds(GI).GetExpActual
        nivel = guilds(GI).GetNivelDeClan
        ExpNecesaria = GetRequiredExpForGuildLevel(nivel)
        If nivel >= 6 Then
            Exit Sub
        End If
        Dim MemberIndex As Byte
        MemberIndex = modGuilds.m_Iterador_ProximoUserIndex(.GuildIndex)
        While MemberIndex > 0
            If UserList(MemberIndex).ConnectionDetails.ConnIDValida Then
                If UserList(MemberIndex).ChatCombate = 1 Then
                    Call SendData(SendTarget.ToIndex, MemberIndex, PrepareMessageLocaleMsg(1789, ExpDar, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1789=Clan> El clan ha ganado ¬1 puntos de experiencia.
                End If
            End If
            MemberIndex = modGuilds.m_Iterador_ProximoUserIndex(.GuildIndex)
        Wend
        ExpActual = ExpActual + ExpDar
        If ExpActual >= ExpNecesaria Then
            'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
            'nivel
            If nivel >= 6 Then
                ExpActual = 0
                ExpNecesaria = 0
                Exit Sub
            End If
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(SND_NIVEL, NO_3D_SOUND, NO_3D_SOUND))
            ExpActual = ExpActual - ExpNecesaria
            nivel = nivel + 1
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1790, nivel, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1790=Clan> El clan ha subido a nivel ¬1. Nuevos beneficios disponibles.
            If nivel > 5 Then
                ExpActual = 0
            End If
        End If
    End With
    guilds(GI).SetExpActual (ExpActual)
    guilds(GI).SetNivelDeClan (nivel)
    Exit Sub
CheckClanExp_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.CheckClanExp", Erl)
End Sub

Public Function GetRequiredExpForGuildLevel(ByVal CurrentLevel As Integer) As Long
    If CurrentLevel = 1 Then
        GetRequiredExpForGuildLevel = 1000
    ElseIf CurrentLevel = 2 Then
        GetRequiredExpForGuildLevel = 2000
    ElseIf CurrentLevel = 3 Then
        GetRequiredExpForGuildLevel = 4000
    ElseIf CurrentLevel = 4 Then
        GetRequiredExpForGuildLevel = 8000
    ElseIf CurrentLevel = 5 Then
        GetRequiredExpForGuildLevel = 16000
    Else
        GetRequiredExpForGuildLevel = 0
    End If
End Function

Public Function MiembrosPermite(ByVal GI As Integer) As Byte
    On Error GoTo MiembrosPermite_Err
    Dim nivel As Byte
    nivel = guilds(GI).GetNivelDeClan
    Select Case nivel
        Case 1
            MiembrosPermite = 5 ' 5 miembros
        Case 2
            MiembrosPermite = 8 ' 3 miembros + pedir ayuda
        Case 3
            MiembrosPermite = 11 ' 3 miembros + seguro de clan
        Case 4
            MiembrosPermite = 14 ' 3 miembros
        Case 5
            MiembrosPermite = 17 ' 3 miembros + barra de vida y de mana
        Case 6
            MiembrosPermite = 20 ' 3 miembros + verse invisible
    End Select
    Exit Function
MiembrosPermite_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.MiembrosPermite", Erl)
End Function

Public Function GetUserGuildMember(ByVal username As String) As String
    On Error GoTo GetUserGuildMember_Err
    GetUserGuildMember = GetUserGuildMemberDatabase(username)
    Exit Function
GetUserGuildMember_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GetUserGuildMember", Erl)
End Function

Public Function GetUserGuildAspirant(ByVal username As String) As Integer
    On Error GoTo GetUserGuildAspirant_Err
    GetUserGuildAspirant = GetUserGuildAspirantDatabase(username)
    Exit Function
GetUserGuildAspirant_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GetUserGuildAspirant", Erl)
End Function

Public Function GetUserGuildPedidos(ByVal username As String) As String
    On Error GoTo GetUserGuildPedidos_Err
    GetUserGuildPedidos = GetUserGuildPedidosDatabase(username)
    Exit Function
GetUserGuildPedidos_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GetUserGuildPedidos", Erl)
End Function

Public Sub SaveUserGuildRejectionReason(ByVal username As String, ByVal Reason As String)
    On Error GoTo SaveUserGuildRejectionReason_Err
    Call SaveUserGuildRejectionReasonDatabase(username, Reason)
    Exit Sub
SaveUserGuildRejectionReason_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildRejectionReason", Erl)
End Sub

Public Sub SaveUserGuildIndex(ByVal UserId As Long, ByVal GuildIndex As Integer)
    On Error GoTo SaveUserGuildIndex_Err
    Call SaveUserGuildIndexDatabase(UserId, GuildIndex)
    Exit Sub
SaveUserGuildIndex_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildIndex", Erl)
End Sub

Public Sub SaveUserGuildAspirant(ByVal UserId As Long, ByVal AspirantIndex As Integer)
    On Error GoTo SaveUserGuildAspirant_Err
    Call SaveUserGuildAspirantDatabase(UserId, AspirantIndex)
    Exit Sub
SaveUserGuildAspirant_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildAspirant", Erl)
End Sub

Public Sub SaveUserGuildMember(ByVal UserId As Long, ByVal guilds As String)
    On Error GoTo SaveUserGuildMember_Err
    Call SaveUserGuildMemberDatabase(UserId, guilds)
    Exit Sub
SaveUserGuildMember_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildMember", Erl)
End Sub

Public Sub SaveUserGuildPedidos(ByVal username As String, ByVal Pedidos As String)
    On Error GoTo SaveUserGuildPedidos_Err
    Call SaveUserGuildPedidosDatabase(username, Pedidos)
    Exit Sub
SaveUserGuildPedidos_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildPedidos", Erl)
End Sub

Public Function GetGuildMemberList(ByVal GuildName As String) As Long()
    On Error GoTo GetGuildMemberList_Err
    Dim i As Integer
    GuildName = UCase$(GuildName)
    For i = LBound(guilds) To UBound(guilds)
        If UCase$(guilds(i).GuildName) = GuildName Then
            GetGuildMemberList = guilds(i).GetMemberList()
            Exit Function
        End If
    Next i
    Dim EmptyList(0) As Long
    GetGuildMemberList = EmptyList
    Exit Function
GetGuildMemberList_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.GetGuildMemberList", Erl)
End Function
