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

'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION      As Byte = 5

'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Public Enum e_ALINEACION_GUILD
    ALINEACION_NEUTRAL = 0
    ALINEACION_ARMADA = 1
    ALINEACION_CAOTICA = 2
    ALINEACION_CIUDADANA = 3
    ALINEACION_CRIMINAL = 4
End Enum

'alineaciones permitidas

Public Enum e_SONIDOS_GUILD

    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45

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
        
        Dim RS As Recordset
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
122     Call TraceError(Err.Number, Err.Description, "modGuilds.LoadGuildsDB", Erl)
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
        
        On Error GoTo m_ConectarMiembroAClan_Err
        

100     If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
102     If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
104         Call guilds(GuildIndex).ConectarMiembro(UserIndex)
106         UserList(UserIndex).GuildIndex = GuildIndex
108         m_ConectarMiembroAClan = True
        End If
        Exit Function
m_ConectarMiembroAClan_Err:
110     Call TraceError(Err.Number, Err.Description, "modGuilds.m_ConectarMiembroAClan", Erl)
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
        
        On Error GoTo m_DesconectarMiembroDelClan_Err
        

100     If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
102     Call guilds(GuildIndex).DesConectarMiembro(UserIndex)

        
        Exit Sub

m_DesconectarMiembroDelClan_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.m_DesconectarMiembroDelClan", Erl)

        
End Sub

Private Function m_EsGuildLeader(ByRef UserId As Long, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo m_EsGuildLeader_Err
100     m_EsGuildLeader = (UserId = guilds(GuildIndex).GetLeader)
        Exit Function
m_EsGuildLeader_Err:
102     Call TraceError(Err.Number, Err.Description, "modGuilds.m_EsGuildLeader", Erl)
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
        
        On Error GoTo m_EsGuildFounder_Err
        
100     m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(GetUserName(guilds(GuildIndex).Fundador))))

        
        Exit Function

m_EsGuildFounder_Err:
102     Call TraceError(Err.Number, Err.Description, "modGuilds.m_EsGuildFounder", Erl)

        
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal ExpellUserId As Long) As Integer
    On Error GoTo m_EcharMiembroDeClan_Err

        'UI echa a Expulsado del clan de Expulsado
        Dim UserReference As t_UserReference
        Dim GI        As Integer
        Dim Map       As Integer
        Dim ExpelledName As String
100     m_EcharMiembroDeClan = 0
        ExpelledName = GetUserName(ExpellUserId)
102     UserReference = NameIndex(ExpelledName)

104     If IsValidUserRef(UserReference) Then
            'pj online
106         GI = UserList(UserReference.ArrayIndex).GuildIndex

108         If GI > 0 Then
110             If m_PuedeSalirDeClan(ExpellUserId, GI, Expulsador) Then
112                 If m_EsGuildLeader(ExpellUserId, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
114                 Call guilds(GI).DesConectarMiembro(UserReference.ArrayIndex)
116                 Call guilds(GI).ExpulsarMiembro(ExpellUserId)
118                 Call LogClanes(ExpelledName & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
120                 UserList(UserReference.ArrayIndex).GuildIndex = 0
122                 map = UserList(UserReference.ArrayIndex).pos.map
124                 If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
126                     Call WriteConsoleMsg(UserReference.ArrayIndex, "Necesitas un clan para pertenecer en este mapa.", e_FontTypeNames.FONTTYPE_INFO)
128                     Call WarpUserChar(UserReference.ArrayIndex, MapInfo(map).Salida.map, MapInfo(map).Salida.x, MapInfo(map).Salida.y, True)
                    Else
130                     Call RefreshCharStatus(UserReference.ArrayIndex)
                    End If

132                 m_EcharMiembroDeClan = GI
                Else
134                 m_EcharMiembroDeClan = 0
                End If
            Else
136             m_EcharMiembroDeClan = 0
            End If

        Else
            'pj offline
140         GI = GetUserGuildIndexDatabase(ExpellUserId)
144         If GI > 0 Then
146             If m_PuedeSalirDeClan(ExpellUserId, GI, Expulsador) Then
148                 If m_EsGuildLeader(ExpellUserId, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
150                 Call guilds(GI).ExpulsarMiembro(ExpellUserId)
152                 Call LogClanes(ExpelledName & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
154                 Map = GetMapDatabase(ExpelledName)
156                 If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
158                     Call SetPositionDatabase(ExpelledName, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.x, MapInfo(Map).Salida.y)
                    End If
160                 m_EcharMiembroDeClan = GI
                Else
162                 m_EcharMiembroDeClan = 0
                End If
            Else
164             m_EcharMiembroDeClan = 0
            End If
        End If
        Exit Function
m_EcharMiembroDeClan_Err:
166     Call TraceError(Err.Number, Err.Description, "modGuilds.m_EcharMiembroDeClan", Erl)
End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
        
        On Error GoTo ActualizarWebSite_Err
        

        Dim GI As Integer

100     GI = UserList(UserIndex).GuildIndex

102     If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
104     If Not m_EsGuildLeader(UserList(UserIndex).id, GI) Then Exit Sub
    
106     Call guilds(GI).SetURL(Web)
    
        
        Exit Sub

ActualizarWebSite_Err:
108     Call TraceError(Err.Number, Err.Description, "modGuilds.ActualizarWebSite", Erl)

        
End Sub

Public Sub ChangeCodexAndDesc(ByRef Desc As String, ByVal GuildIndex As Integer)
        
        On Error GoTo ChangeCodexAndDesc_Err
        

        Dim i As Long
    
100     If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    
102     With guilds(GuildIndex)
104         Call .SetDesc(Desc)

        End With

        
        Exit Sub

ChangeCodexAndDesc_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.ChangeCodexAndDesc", Erl)

        
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
        
        On Error GoTo ActualizarNoticias_Err
        

        Dim GI As Integer

100     GI = UserList(UserIndex).GuildIndex
    
102     If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
104     If Not m_EsGuildLeader(UserList(UserIndex).id, GI) Then Exit Sub
    
106     Call guilds(GI).SetGuildNews(Datos)
        
        
        Exit Sub

ActualizarNoticias_Err:
108     Call TraceError(Err.Number, Err.Description, "modGuilds.ActualizarNoticias", Erl)

        
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef Desc As String, ByRef GuildName As String, ByVal Alineacion As e_ALINEACION_GUILD, ByRef refError As String) As Boolean
        
        On Error GoTo CrearNuevoClan_Err
        

        Dim i           As Integer

        Dim DummyString As String

100     CrearNuevoClan = False

102     If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
104         refError = DummyString
            Exit Function

        End If

106     If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
108         refError = "Nombre de clan inválido."
            Exit Function

        End If
    
110     If YaExiste(GuildName) Then
112         refError = "Ya existe un clan con ese nombre."
            Exit Function

        End If

        'tenemos todo para fundar ya
114     If CANTIDADDECLANES < UBound(guilds) Then
116         CANTIDADDECLANES = CANTIDADDECLANES + 1
            'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

            'constructor custom de la clase clan
118         Set guilds(CANTIDADDECLANES) = New clsClan
        
            'Damos de alta al clan como nuevo inicializando sus archivos
122         Call guilds(CANTIDADDECLANES).InicializarNuevoClan(GuildName, CANTIDADDECLANES, Alineacion, UserList(FundadorIndex).id)
        
            'seteamos codex y descripcion
124         Call guilds(CANTIDADDECLANES).SetDesc(Desc)
126         Call guilds(CANTIDADDECLANES).SetGuildNews("¡Bienvenido a " & GuildName & "! Clan creado con alineación : " & Alineacion2String(Alineacion) & ".")
128         Call guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).id)
        
130         Call guilds(CANTIDADDECLANES).SetNivelDeClan(1)
        
132         Call guilds(CANTIDADDECLANES).SetExpActual(0)
        
            '"conectamos" al nuevo miembro a la lista de la clase
136         Call guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).id)
138         Call guilds(CANTIDADDECLANES).ConectarMiembro(FundadorIndex)
140         UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
142         Call RefreshCharStatus(FundadorIndex)
        
144         For i = 1 To CANTIDADDECLANES - 1
146             Call guilds(i).ProcesarFundacionDeOtroClan
148         Next i

        Else
150         refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
            Exit Function

        End If
    
152     CrearNuevoClan = True

        
        Exit Function

CrearNuevoClan_Err:
154     Call TraceError(Err.Number, Err.Description, "modGuilds.CrearNuevoClan", Erl)

        
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer, ByRef guildList() As String)
        
        On Error GoTo SendGuildNews_Err
        

        Dim GuildIndex As Integer

        Dim i          As Integer

        Dim go         As Integer

        Dim ClanNivel  As Byte

        Dim ExpAcu     As Integer

        Dim ExpNe      As Integer

100     GuildIndex = UserList(UserIndex).GuildIndex

102     If GuildIndex = 0 Then Exit Sub
    
        Dim MemberList() As Long
104     MemberList = guilds(GuildIndex).GetMemberList()
106     ClanNivel = guilds(GuildIndex).GetNivelDeClan
108     ExpAcu = guilds(GuildIndex).GetExpActual
110     ExpNe = GetRequiredExpForGuildLevel(ClanNivel)
112     Call WriteGuildNews(UserIndex, guilds(GuildIndex).GetGuildNews, guildList, MemberList, ClanNivel, ExpAcu, ExpNe)
        Exit Sub

SendGuildNews_Err:
114     Call TraceError(Err.Number, Err.Description, "modGuilds.SendGuildNews", Erl)

        
End Sub

Public Function m_PuedeSalirDeClan(ByRef UserId As Long, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
        'sale solo si no es fundador del clan.
        On Error GoTo m_PuedeSalirDeClan_Err

100     m_PuedeSalirDeClan = False
102     If GuildIndex = 0 Or guilds(GuildIndex) Is Nothing Then Exit Function
    
        'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
104     If QuienLoEchaUI = -1 Then
106         m_PuedeSalirDeClan = True
            Exit Function
        End If

        'cuando UI no puede echar a nombre?
        'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
108     If UserList(QuienLoEchaUI).flags.Privilegios And e_PlayerType.user Then
110         If Not m_EsGuildLeader(UserList(QuienLoEchaUI).id, GuildIndex) Then
112             If UserList(QuienLoEchaUI).id <> UserId Then      'si no sale voluntariamente...
                    Exit Function
                End If
            End If
        End If
114     m_PuedeSalirDeClan = guilds(GuildIndex).Fundador <> UserId
        Exit Function
m_PuedeSalirDeClan_Err:
116     Call TraceError(Err.Number, Err.Description, "modGuilds.m_PuedeSalirDeClan", Erl)
End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As e_ALINEACION_GUILD, ByRef refError As String) As Boolean
        
        On Error GoTo PuedeFundarUnClan_Err
        

100     PuedeFundarUnClan = False

102     If UserList(UserIndex).GuildIndex > 0 Then
104         refError = "Ya perteneces a un clan, no podés fundar otro"
            Exit Function
        End If
    
106     If UserList(UserIndex).Stats.ELV < 23 Or UserList(UserIndex).Stats.UserSkills(e_Skill.liderazgo) < 50 Then
108         refError = "Para fundar un clan debes ser Nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar."
            Exit Function
        End If
    
110     If Not TieneObjetos(407, 1, UserIndex) Then
112         refError = "Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar."
            Exit Function
        End If
    
114     If Not TieneObjetos(408, 1, UserIndex) Then
116         refError = "Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar."
            Exit Function
        End If
        
121             If Not TieneObjetos(409, 1, UserIndex) Then
122         refError = "Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar."
            Exit Function
        End If
    
123     If Not TieneObjetos(412, 1, UserIndex) Then
124         refError = "Para fundar un clan debes ser nivel 23, tener 50 puntos en liderazgo y tener en tu inventario las Gemas de Fundación Verde, Roja, Azul y Polar."
            Exit Function
        End If
        
        If Alineacion = e_ALINEACION_GUILD.ALINEACION_CIUDADANA And UserList(UserIndex).flags.Seguro = False Then
            refError = "Para fundar un clan ciudadano deberás tener activado el seguro."
            Exit Function
        End If
    
125     Select Case Alineacion
            Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
126             If status(UserIndex) = e_Facciones.Caos Or status(UserIndex) = e_Facciones.Armada Or status(UserIndex) = e_Facciones.consejo Or status(UserIndex) = e_Facciones.concilio Then
127                 refError = "Para fundar un clan neutral deberás ser ciudadano o criminal."
                    Exit Function
                End If

128         Case e_ALINEACION_GUILD.ALINEACION_ARMADA

129             If status(UserIndex) <> e_Facciones.Armada And status(UserIndex) <> e_Facciones.consejo Then
130                 refError = "Para fundar un clan de la Armada Real deberás pertenecer a la misma."
                    Exit Function
                End If
                
131         Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
132             If status(UserIndex) <> e_Facciones.Caos And status(UserIndex) <> e_Facciones.concilio Then
133                 refError = "Para fundar un clan de la Legión Oscura deberás pertenecer a la misma."
                    Exit Function
                End If
                
            Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
             If status(UserIndex) <> e_Facciones.Ciudadano And status(UserIndex) <> e_Facciones.Armada Then
                refError = "Para fundar un clan ciudadano deberás ser ciudadano."
                Exit Function
            End If
                
            Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
             If Status(UserIndex) <> e_Facciones.Criminal And Status(UserIndex) <> e_Facciones.Caos Then
                refError = "Para fundar un clan criminal deberás ser criminal o legión oscura."
                Exit Function
            End If
            
        End Select

136     PuedeFundarUnClan = True
    
        
        Exit Function

PuedeFundarUnClan_Err:
138     Call TraceError(Err.Number, Err.Description, "modGuilds.PuedeFundarUnClan", Erl)

        
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
        
        On Error GoTo m_EstadoPermiteEntrarChar_Err
        

        Dim Promedio As Long

        Dim ELV      As Integer

        Dim f        As Byte

100     m_EstadoPermiteEntrarChar = False
    
102     If InStrB(Personaje, "\") <> 0 Then
104         Personaje = Replace(Personaje, "\", vbNullString)

        End If

106     If InStrB(Personaje, "/") <> 0 Then
108         Personaje = Replace(Personaje, "/", vbNullString)

        End If

110     If InStrB(Personaje, ".") <> 0 Then
112         Personaje = Replace(Personaje, ".", vbNullString)

        End If
    
114     If PersonajeExiste(Personaje) Then
            Dim status As Integer
            status = CInt(GetUserValue(LCase$(Personaje), "status"))
        
118         Select Case guilds(GuildIndex).Alineacion

                Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
120                 m_EstadoPermiteEntrarChar = (status = e_Facciones.Ciudadano Or status = e_Facciones.Criminal)

122             Case e_ALINEACION_GUILD.ALINEACION_ARMADA
124                 m_EstadoPermiteEntrarChar = (status = e_Facciones.Armada Or status = e_Facciones.consejo)

126             Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
128                 m_EstadoPermiteEntrarChar = (status = e_Facciones.Caos Or status = e_Facciones.concilio)
                
                Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
                     m_EstadoPermiteEntrarChar = (status = e_Facciones.Ciudadano Or status = e_Facciones.Armada)
                
                Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
                     m_EstadoPermiteEntrarChar = (status = e_Facciones.Criminal Or status = e_Facciones.Caos)

            End Select

        End If

        
        Exit Function

m_EstadoPermiteEntrarChar_Err:
130     Call TraceError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrarChar", Erl)

        
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
        On Error GoTo m_EstadoPermiteEntrar_Err

100     Select Case guilds(GuildIndex).Alineacion

            Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
102           m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Ciudadano Or Status(UserIndex) = e_Facciones.Criminal

104         Case e_ALINEACION_GUILD.ALINEACION_ARMADA
106           m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo

108         Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
110           m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio

            Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
               m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Ciudadano Or Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo
               
            Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrar = Status(UserIndex) = e_Facciones.Criminal Or Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio
                
        End Select

        Exit Function

m_EstadoPermiteEntrar_Err:
112     Call TraceError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrar", Erl)


End Function

Public Function String2Alineacion(ByRef S As String) As e_ALINEACION_GUILD
        
        On Error GoTo String2Alineacion_Err

100     Select Case S

            Case "Neutral"
102             String2Alineacion = e_ALINEACION_GUILD.ALINEACION_NEUTRAL
104         Case "Armada Real"
106             String2Alineacion = e_ALINEACION_GUILD.ALINEACION_ARMADA
108         Case "Legión Oscura"
110             String2Alineacion = e_ALINEACION_GUILD.ALINEACION_CAOTICA
            Case "Ciudadano"
                String2Alineacion = e_ALINEACION_GUILD.ALINEACION_CIUDADANA
            Case "Criminal"
                String2Alineacion = e_ALINEACION_GUILD.ALINEACION_CRIMINAL

        End Select

        
        Exit Function

String2Alineacion_Err:
112     Call TraceError(Err.Number, Err.Description, "modGuilds.String2Alineacion", Erl)

        
End Function

Public Function Alineacion2String(ByVal Alineacion As e_ALINEACION_GUILD) As String
        
        On Error GoTo Alineacion2String_Err
        

100     Select Case Alineacion

            Case e_ALINEACION_GUILD.ALINEACION_NEUTRAL
102             Alineacion2String = "Neutral"

104         Case e_ALINEACION_GUILD.ALINEACION_ARMADA
106             Alineacion2String = "Armada Real"

108         Case e_ALINEACION_GUILD.ALINEACION_CAOTICA
110             Alineacion2String = "Legión Oscura"

            Case e_ALINEACION_GUILD.ALINEACION_CIUDADANA
                Alineacion2String = "Ciudadano"

            Case e_ALINEACION_GUILD.ALINEACION_CRIMINAL
                 Alineacion2String = "Criminal"
        End Select

        
        Exit Function

Alineacion2String_Err:
112     Call TraceError(Err.Number, Err.Description, "modGuilds.Alineacion2String", Erl)

        
End Function

Public Function Relacion2String(ByVal Relacion As e_RELACIONES_GUILD) As String
        
        On Error GoTo Relacion2String_Err
        

100     Select Case Relacion

            Case e_RELACIONES_GUILD.ALIADOS
102             Relacion2String = "A"

104         Case e_RELACIONES_GUILD.GUERRA
106             Relacion2String = "G"

108         Case e_RELACIONES_GUILD.PAZ
110             Relacion2String = "P"

112         Case Else
114             Relacion2String = "?"

        End Select

        
        Exit Function

Relacion2String_Err:
116     Call TraceError(Err.Number, Err.Description, "modGuilds.Relacion2String", Erl)

        
End Function

Public Function String2Relacion(ByVal S As String) As e_RELACIONES_GUILD
        
        On Error GoTo String2Relacion_Err
        

100     Select Case UCase$(Trim$(S))

            Case vbNullString, "P"
102             String2Relacion = e_RELACIONES_GUILD.PAZ

104         Case "G"
106             String2Relacion = e_RELACIONES_GUILD.GUERRA

108         Case "A"
110             String2Relacion = e_RELACIONES_GUILD.ALIADOS

112         Case Else
114             String2Relacion = e_RELACIONES_GUILD.PAZ

        End Select

        
        Exit Function

String2Relacion_Err:
116     Call TraceError(Err.Number, Err.Description, "modGuilds.String2Relacion", Erl)

        
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
        
        On Error GoTo GuildNameValido_Err
        

        Dim car As Byte

        Dim i   As Integer

        'old function by morgo

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))

106         If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
108             GuildNameValido = False
                Exit Function

            End If
    
110     Next i

112     GuildNameValido = True

        
        Exit Function

GuildNameValido_Err:
114     Call TraceError(Err.Number, Err.Description, "modGuilds.GuildNameValido", Erl)

        
End Function

Public Function YaExiste(ByVal GuildName As String) As Boolean
    On Error GoTo YaExiste_Err
        Dim i As Integer
100     YaExiste = False
102     GuildName = UCase$(GuildName)
104     For i = 1 To CANTIDADDECLANES
106         YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
108         If YaExiste Then Exit Function
110     Next i
        Exit Function
YaExiste_Err:
112     Call TraceError(Err.Number, Err.Description, "modGuilds.YaExiste", Erl)
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
        On Error GoTo GuildIndex_Err

        'me da el indice del guildname
        Dim i As Integer

100     GuildIndex = 0
102     GuildName = UCase$(GuildName)

104     For i = 1 To CANTIDADDECLANES
106         If UCase$(guilds(i).GuildName) = GuildName Then
108             GuildIndex = i
                Exit Function
            End If
110     Next i
        Exit Function

GuildIndex_Err:
112     Call TraceError(Err.Number, Err.Description, "modGuilds.GuildIndex", Erl)
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
        
        On Error GoTo m_ListaDeMiembrosOnline_Err
        

        Dim i As Integer
    
100     If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
102         i = guilds(GuildIndex).m_Iterador_ProximoUserIndex

104         While i > 0

                'No mostramos dioses y admins
106             If i <> UserIndex And ((UserList(i).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) <> 0 Or (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) <> 0)) Then m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
108             i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
            Wend

        End If

110     If Len(m_ListaDeMiembrosOnline) > 0 Then
112         m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)

        End If

        
        Exit Function

m_ListaDeMiembrosOnline_Err:
114     Call TraceError(Err.Number, Err.Description, "modGuilds.m_ListaDeMiembrosOnline", Erl)

        
End Function

Public Function PrepareGuildsList() As String()
        
        On Error GoTo PrepareGuildsList_Err
        

        Dim tStr() As String

        Dim i      As Long
    
100     If CANTIDADDECLANES = 0 Then
102         ReDim tStr(0) As String
        Else
104         ReDim tStr(CANTIDADDECLANES - 1) As String
        
106         For i = 1 To CANTIDADDECLANES
108             tStr(i - 1) = guilds(i).GuildName & "-" & guilds(i).Alineacion
110         Next i

        End If
    
112     PrepareGuildsList = tStr

        
        Exit Function

PrepareGuildsList_Err:
114     Call TraceError(Err.Number, Err.Description, "modGuilds.PrepareGuildsList", Erl)

        
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
        
        On Error GoTo SendGuildDetails_Err
        

        Dim codex(CANTIDADMAXIMACODEX - 1) As String

        Dim GI                             As Integer

        Dim i                              As Long

100     GI = GuildIndex(GuildName)

102     If GI = 0 Then Exit Sub
    
104     With guilds(GI)
106         Call WriteGuildDetails(UserIndex, GuildName, GetUserName(.Fundador), .GetFechaFundacion, .GetLeader, .CantidadDeMiembros, Alineacion2String(.Alineacion), .GetDesc, .GetNivelDeClan)

        End With

        
        Exit Sub

SendGuildDetails_Err:
108     Call TraceError(Err.Number, Err.Description, "modGuilds.SendGuildDetails", Erl)

        
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
    On Error GoTo SendGuildLeaderInfo_Err
    
        Dim GI              As Integer
        Dim guildList()     As String
        Dim MemberList()    As Long
        Dim aspirantsList() As String
    
100     With UserList(UserIndex)
102         GI = .GuildIndex
104         guildList = PrepareGuildsList()

106         If GI <= 0 Or GI > CANTIDADDECLANES Then
                'Send the guild list instead
108             Call WriteGuildList(UserIndex, guildList)
                Exit Sub
            End If
        
110         If Not m_EsGuildLeader(.id, GI) Then
                'Send the guild list instead
112             Call modGuilds.SendGuildNews(UserIndex, guildList)
                Exit Sub
            End If
        
114         MemberList = guilds(GI).GetMemberList()
116         aspirantsList = guilds(GI).GetAspirantes()
            Dim ClanLevel As Integer
            ClanLevel = guilds(GI).GetNivelDeClan
118         Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList, ClanLevel, guilds(GI).GetExpActual, GetRequiredExpForGuildLevel(ClanLevel))

        End With

        
        Exit Sub

SendGuildLeaderInfo_Err:
120     Call TraceError(Err.Number, Err.Description, "modGuilds.SendGuildLeaderInfo", Erl)

        
End Sub

Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
        'itera sobre los onlinemembers
        
        On Error GoTo m_Iterador_ProximoUserIndex_Err
        
100     m_Iterador_ProximoUserIndex = 0

102     If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
104         m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()

        End If

        
        Exit Function

m_Iterador_ProximoUserIndex_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.m_Iterador_ProximoUserIndex", Erl)

        
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
        'itera sobre los gms escuchando este clan
        
        On Error GoTo Iterador_ProximoGM_Err
        
100     Iterador_ProximoGM = 0

102     If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
104         Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()

        End If

        
        Exit Function

Iterador_ProximoGM_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.Iterador_ProximoGM", Erl)
End Function

Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
        
        On Error GoTo GMEscuchaClan_Err
        

        Dim GI As Integer

        'listen to no guild at all
100     If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
            'Quit listening to previous guild!!
102         Call WriteLocaleMsg(UserIndex, 1603, guilds(UserList(UserIndex).EscucheClan).GuildName, e_FontTypeNames.FONTTYPE_GUILD) 'Msg1603= Dejas de escuchar a : ¬1
104         guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            Exit Function

        End If
    
        'devuelve el guildindex
106     GI = GuildIndex(GuildName)

108     If GI > 0 Then
110         If UserList(UserIndex).EscucheClan <> 0 Then
112             If UserList(UserIndex).EscucheClan = GI Then
                    'Already listening to them...
114                 Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, e_FontTypeNames.FONTTYPE_GUILD)
116                 GMEscuchaClan = GI
                    Exit Function
                Else
                    'Quit listening to previous guild!!
118                 Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, e_FontTypeNames.FONTTYPE_GUILD)
120                 guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)

                End If

            End If
        
122         Call guilds(GI).ConectarGM(UserIndex)
124         Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, e_FontTypeNames.FONTTYPE_GUILD)
126         GMEscuchaClan = GI
128         UserList(UserIndex).EscucheClan = GI
        Else
130         Call WriteConsoleMsg(UserIndex, "Error, el clan no existe", e_FontTypeNames.FONTTYPE_GUILD)
132         GMEscuchaClan = 0

        End If
    
        
        Exit Function

GMEscuchaClan_Err:
134     Call TraceError(Err.Number, Err.Description, "modGuilds.GMEscuchaClan", Erl)

        
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
        'el index lo tengo que tener de cuando me puse a escuchar
        
        On Error GoTo GMDejaDeEscucharClan_Err
        
100     UserList(UserIndex).EscucheClan = 0
102     Call guilds(GuildIndex).DesconectarGM(UserIndex)

        
        Exit Sub

GMDejaDeEscucharClan_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.GMDejaDeEscucharClan", Erl)

        
End Sub

Public Function PersonajeEsLeader(ByVal CharId As Long) As Boolean
        Dim GuildIndex As Integer
100     GuildIndex = GetUserGuildIndexDatabase(CharId)
102     If GuildIndex > 0 Then
104         If m_EsGuildLeader(CharId, GuildIndex) Then PersonajeEsLeader = True
        End If
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
        
        On Error GoTo a_RechazarAspiranteChar_Err
        

100     If InStrB(Aspirante, "\") <> 0 Then
102         Aspirante = Replace(Aspirante, "\", "")

        End If

104     If InStrB(Aspirante, "/") <> 0 Then
106         Aspirante = Replace(Aspirante, "/", "")

        End If

108     If InStrB(Aspirante, ".") <> 0 Then
110         Aspirante = Replace(Aspirante, ".", "")

        End If

112     Call SaveUserGuildRejectionReason(Aspirante, Detalles)

        
        Exit Sub

a_RechazarAspiranteChar_Err:
114     Call TraceError(Err.Number, Err.Description, "modGuilds.a_RechazarAspiranteChar", Erl)

        
End Sub

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef nombre As String, ByRef refError As String) As Boolean
        
        On Error GoTo a_RechazarAspirante_Err

        Dim GI           As Integer
        Dim NroAspirante As Integer

100     a_RechazarAspirante = False
102     GI = UserList(UserIndex).GuildIndex

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No perteneces a ningún clan"
            Exit Function
        End If

114     Call guilds(GI).RetirarAspirante(nombre)
116     refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
118     a_RechazarAspirante = True
        
        Exit Function

a_RechazarAspirante_Err:
120     Call TraceError(Err.Number, Err.Description, "modGuilds.a_RechazarAspirante", Erl)
End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef name As String) As String
        
        On Error GoTo a_DetallesAspirante_Err

        Dim GI           As Integer
        Dim NroAspirante As Integer
100     GI = UserList(UserIndex).GuildIndex
102     If GI <= 0 Or GI > CANTIDADDECLANES Then
            Exit Function
        End If
    
104     If Not m_EsGuildLeader(UserList(UserIndex).id, GI) Then
            Exit Function
        End If
    
106     a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(name)
    
        
        Exit Function

a_DetallesAspirante_Err:
112     Call TraceError(Err.Number, Err.Description, "modGuilds.a_DetallesAspirante", Erl)

        
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)

        Dim GI     As Integer
        Dim NroAsp As Integer
        Dim list() As Long
        Dim i      As Long
        On Error GoTo Error

100     GI = UserList(UserIndex).GuildIndex
102     Personaje = UCase$(Personaje)

        If Not PersonajeExiste(Personaje) Then
                Call guilds(GI).ExpulsarMiembro(Personaje)
                Call WriteConsoleMsg(UserIndex, "El personaje no existe y fue eliminado de la lista de miembros", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         Call WriteConsoleMsg(UserIndex, "No perteneces a ningún clan.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).id, GI) Then
110         Call WriteConsoleMsg(UserIndex, "No eres el lider de tu clan.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
112     If InStrB(Personaje, "\") <> 0 Then
114         Personaje = Replace$(Personaje, "\", vbNullString)

        End If

116     If InStrB(Personaje, "/") <> 0 Then
118         Personaje = Replace$(Personaje, "/", vbNullString)

        End If

120     If InStrB(Personaje, ".") <> 0 Then
122         Personaje = Replace$(Personaje, ".", vbNullString)

        End If
        
        Dim HasRequest As Boolean
        Dim CharId As Long
        CharId = GetCharacterIdWithName(Personaje)
        HasRequest = guilds(GI).HasGuildRequest(CharId)
    
126     If Not HasRequest Then
128         list = guilds(GI).GetMemberList()
130         For i = 0 To UBound(list())
132             If CharId = list(i) Then Exit For
134         Next i
136         If i > UBound(list()) Then
138             Call WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro del clan.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
140     Call SendCharacterInfoDatabase(UserIndex, Personaje)
        Exit Sub
Error:
142     If Not PersonajeExiste(Personaje) Then
144         Call LogError("El usuario " & UserList(UserIndex).Name & " (" & UserIndex & " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
        Else
146         Call LogError("[" & Err.Number & "] " & Err.Description & " En la rutina SendDetallesPersonaje, por el usuario " & UserList(UserIndex).Name & " (" & UserIndex & " ), pidiendo informacion sobre el personaje " & Personaje)
        End If
End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
        
        On Error GoTo a_NuevoAspirante_Err
        

        Dim ViejoSolicitado   As String

        Dim ViejoGuildINdex   As Integer

        Dim ViejoNroAspirante As Integer

        Dim NuevoGuildIndex   As Integer

100     a_NuevoAspirante = False

102     If UserList(UserIndex).GuildIndex > 0 Then
104         refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
            Exit Function

        End If
    
106     If EsNewbie(UserIndex) Then
108         refError = "Los newbies no tienen derecho a entrar a un clan."
            Exit Function

        End If

110     NuevoGuildIndex = GuildIndex(clan)

112     If NuevoGuildIndex = 0 Then
114         refError = "Ese clan no existe! Avise a un administrador."
            Exit Function

        End If
    
116     If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
118         refError = "Tu no podés entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
            Exit Function
        End If

120     If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
122         refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
            Exit Function
        End If
        
        Dim NuevoGuildAspirantes() As String
124     NuevoGuildAspirantes = guilds(NuevoGuildIndex).GetAspirantes()

        Dim i As Long
126     For i = 0 To UBound(NuevoGuildAspirantes)
            
128         If UserList(UserIndex).Name = NuevoGuildAspirantes(i) Then
130             refError = "Ya has enviado una solicitud a este clan."
                Exit Function

            End If
                    
        Next

132     ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).Name & ".chr", "GUILD", "ASPIRANTEA")

134     If LenB(ViejoSolicitado) <> 0 Then
            'borramos la vieja solicitud
136         ViejoGuildINdex = CInt(ViejoSolicitado)

138         If ViejoGuildINdex <> 0 Then
140             Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).name)
            End If
        End If
    
146     Call SendData(SendTarget.ToDiosesYclan, NuevoGuildIndex, PrepareMessageGuildChat("Clan: [" & UserList(UserIndex).Name & "] ha enviado solicitud para unirse al clan.", 7))
    
148     Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).Name, Solicitud)
150     a_NuevoAspirante = True
        
        Exit Function

a_NuevoAspirante_Err:
152     Call TraceError(Err.Number, Err.Description, "modGuilds.a_NuevoAspirante", Erl)

        
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
        
        On Error GoTo a_AceptarAspirante_Err
        

        Dim GI           As Integer
        Dim tGI          As Integer
        Dim AspiranteRef  As t_UserReference

        'un pj ingresa al clan :D

100     a_AceptarAspirante = False
    
102     GI = UserList(UserIndex).GuildIndex

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No perteneces a ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).id, GI) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
        Dim UserDidRequest As Boolean
        Dim CharId As Long
        CharId = GetCharacterIdWithName(Aspirante)
        UserDidRequest = guilds(GI).HasGuildRequest(CharId)
    
114     If Not UserDidRequest Then
116         refError = "El Pj no es aspirante al clan"
            Exit Function
        End If
    
118     AspiranteRef = NameIndex(Aspirante)

120     If IsValidUserRef(AspiranteRef) Then
            'pj Online
122         If Not m_EstadoPermiteEntrar(AspiranteRef.ArrayIndex, GI) Then
124             refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
126             Call guilds(GI).RetirarAspirante(Aspirante)
                Exit Function
128         ElseIf Not UserList(AspiranteRef.ArrayIndex).GuildIndex = 0 Then
130             refError = Aspirante & " ya es parte de otro clan."
132             Call guilds(GI).RetirarAspirante(Aspirante)
                Exit Function
            End If
            
            If GuildAlignmentIndex(GI) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA And UserList(AspiranteRef.ArrayIndex).flags.Seguro = False Then
                refError = Aspirante & " deberá activar el seguro para entrar al clan."
                Call guilds(GI).RetirarAspirante(Aspirante)
               Exit Function
            End If
        Else
134         If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
136             refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
138             Call guilds(GI).RetirarAspirante(Aspirante)
                Exit Function
            Else
142             tGI = GetUserGuildIndexDatabase(CharId)
146             If tGI <> 0 Then
148                 refError = Aspirante & " ya es parte de otro clan."
150                 Call guilds(GI).RetirarAspirante(Aspirante)
                    Exit Function
                End If
            End If
        End If
    
152     If guilds(GI).CantidadDeMiembros >= MiembrosPermite(GI) Then
154         refError = "La capacidad del clan esta completa."
            Exit Function

        End If
    
        'el pj es aspirante al clan y puede entrar
    
156     Call guilds(GI).RetirarAspirante(Aspirante)
158     Call guilds(GI).AceptarNuevoMiembro(CharId)
    
        ' If player is online, update tag
160     If IsValidUserRef(AspiranteRef) Then
162         Call RefreshCharStatus(AspiranteRef.ArrayIndex)

        End If
    
164     a_AceptarAspirante = True

        
        Exit Function

a_AceptarAspirante_Err:
166     Call TraceError(Err.Number, Err.Description, "modGuilds.a_AceptarAspirante", Erl)

        
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
        
        On Error GoTo GuildName_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     GuildName = guilds(GuildIndex).GuildName

        
        Exit Function

GuildName_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.GuildName", Erl)

        
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
        
        On Error GoTo GuildLeader_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
102     GuildLeader = guilds(GuildIndex).GetLeader

        
        Exit Function

GuildLeader_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.GuildLeader", Erl)

        
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
        
        On Error GoTo GuildAlignment_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)

        
        Exit Function

GuildAlignment_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.GuildAlignment", Erl)

        
End Function


Public Function GuildAlignmentIndex(ByVal GuildIndex As Integer) As e_ALINEACION_GUILD
        
        On Error GoTo GuildAlignmentIndex_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     GuildAlignmentIndex = guilds(GuildIndex).Alineacion
        
        Exit Function

GuildAlignmentIndex_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.GuildAlignmentIndex", Erl)

        
End Function

Public Function NivelDeClan(ByVal GuildIndex As Integer) As Byte
        
        On Error GoTo NivelDeClan_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     NivelDeClan = guilds(GuildIndex).GetNivelDeClan

        
        Exit Function

NivelDeClan_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.NivelDeClan", Erl)

        
End Function

Public Function Alineacion(ByVal GuildIndex As Integer) As Byte
        
        On Error GoTo Alineacion_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     Alineacion = guilds(GuildIndex).Alineacion

        
        Exit Function

Alineacion_Err:
104     Call TraceError(Err.Number, Err.Description, "modGuilds.Alineacion", Erl)

        
End Function

Sub CheckClanExp(ByVal UserIndex As Integer, ByVal ExpDar As Integer)
        
        On Error GoTo CheckClanExp_Err
        
        Dim ExpActual    As Integer
        Dim ExpNecesaria As Integer
        Dim GI           As Integer
        Dim nivel        As Byte
        
        With UserList(UserIndex)
100         GI = .GuildIndex
102         ExpActual = guilds(GI).GetExpActual
104         nivel = guilds(GI).GetNivelDeClan
106         ExpNecesaria = GetRequiredExpForGuildLevel(nivel)
    
108         If nivel >= 6 Then
                Exit Sub
            End If

            Dim MemberIndex As Byte
            MemberIndex = modGuilds.m_Iterador_ProximoUserIndex(.GuildIndex)
            
            While MemberIndex > 0
                If UserList(MemberIndex).ConnectionDetails.ConnIDValida Then
                    If UserList(MemberIndex).ChatCombate = 1 Then
                        Call SendData(SendTarget.ToIndex, MemberIndex, PrepareMessageConsoleMsg("Clan> El clan ha ganado " & ExpDar & " puntos de experiencia.", e_FontTypeNames.FONTTYPE_GUILD))
                    End If
                End If
            
                MemberIndex = modGuilds.m_Iterador_ProximoUserIndex(.GuildIndex)
            Wend
114         ExpActual = ExpActual + ExpDar
    
116         If ExpActual >= ExpNecesaria Then
                'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
                'nivel
118             If nivel >= 6 Then
120                 ExpActual = 0
122                 ExpNecesaria = 0
                    Exit Sub
                End If
    
124             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(SND_NIVEL, NO_3D_SOUND, NO_3D_SOUND))
126             ExpActual = ExpActual - ExpNecesaria
128             nivel = nivel + 1
130             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Clan> El clan ha subido a nivel " & nivel & ". Nuevos beneficios disponibles.", e_FontTypeNames.FONTTYPE_GUILD))
        
132             If nivel > 5 Then
146                 ExpActual = 0
                End If
            End If
        End With
150     guilds(GI).SetExpActual (ExpActual)
152     guilds(GI).SetNivelDeClan (nivel)
        Exit Sub
CheckClanExp_Err:
154     Call TraceError(Err.Number, Err.Description, "modGuilds.CheckClanExp", Erl)
        
End Sub

Public Function GetRequiredExpForGuildLevel(ByVal CurrentLevel As Integer) As Long
128         If CurrentLevel = 1 Then
130             GetRequiredExpForGuildLevel = 1000
132         ElseIf CurrentLevel = 2 Then
134             GetRequiredExpForGuildLevel = 2000
136         ElseIf CurrentLevel = 3 Then
138             GetRequiredExpForGuildLevel = 4000
140         ElseIf CurrentLevel = 4 Then
142             GetRequiredExpForGuildLevel = 8000
143         ElseIf CurrentLevel = 5 Then
145             GetRequiredExpForGuildLevel = 16000
            Else
144             GetRequiredExpForGuildLevel = 0
            End If
End Function
Public Function MiembrosPermite(ByVal GI As Integer) As Byte
        
        On Error GoTo MiembrosPermite_Err
        

        Dim nivel As Byte

100     nivel = guilds(GI).GetNivelDeClan

102     Select Case nivel

            Case 1
104             MiembrosPermite = 5 ' 5 miembros

106         Case 2
108             MiembrosPermite = 8 ' 3 miembros + pedir ayuda

110         Case 3
112             MiembrosPermite = 11 ' 3 miembros + seguro de clan

114         Case 4
116             MiembrosPermite = 14 ' 3 miembros

            Case 5
                MiembrosPermite = 17 ' 3 miembros + barra de vida y de mana
                
            Case 6
                MiembrosPermite = 20 ' 3 miembros + verse invisible

        End Select

        
        Exit Function

MiembrosPermite_Err:
118     Call TraceError(Err.Number, Err.Description, "modGuilds.MiembrosPermite", Erl)

        
End Function

Public Function GetUserGuildMember(ByVal UserName As String) As String
        
        On Error GoTo GetUserGuildMember_Err

104     GetUserGuildMember = GetUserGuildMemberDatabase(UserName)

        Exit Function

GetUserGuildMember_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.GetUserGuildMember", Erl)

        
End Function

Public Function GetUserGuildAspirant(ByVal UserName As String) As Integer
        
        On Error GoTo GetUserGuildAspirant_Err

104     GetUserGuildAspirant = GetUserGuildAspirantDatabase(UserName)

        Exit Function

GetUserGuildAspirant_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.GetUserGuildAspirant", Erl)

        
End Function

Public Function GetUserGuildPedidos(ByVal UserName As String) As String
        
        On Error GoTo GetUserGuildPedidos_Err

104     GetUserGuildPedidos = GetUserGuildPedidosDatabase(UserName)

        Exit Function

GetUserGuildPedidos_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.GetUserGuildPedidos", Erl)

        
End Function

Public Sub SaveUserGuildRejectionReason(ByVal UserName As String, ByVal Reason As String)
        
        On Error GoTo SaveUserGuildRejectionReason_Err

104     Call SaveUserGuildRejectionReasonDatabase(UserName, Reason)

        Exit Sub

SaveUserGuildRejectionReason_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildRejectionReason", Erl)

        
End Sub

Public Sub SaveUserGuildIndex(ByVal UserId As Long, ByVal GuildIndex As Integer)
        
        On Error GoTo SaveUserGuildIndex_Err

104     Call SaveUserGuildIndexDatabase(UserId, GuildIndex)

        Exit Sub

SaveUserGuildIndex_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildIndex", Erl)

        
End Sub

Public Sub SaveUserGuildAspirant(ByVal UserId As Long, ByVal AspirantIndex As Integer)
        
        On Error GoTo SaveUserGuildAspirant_Err
        
104     Call SaveUserGuildAspirantDatabase(UserId, AspirantIndex)

        Exit Sub

SaveUserGuildAspirant_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildAspirant", Erl)

        
End Sub

Public Sub SaveUserGuildMember(ByVal UserId As Long, ByVal guilds As String)
    On Error GoTo SaveUserGuildMember_Err
104     Call SaveUserGuildMemberDatabase(UserId, guilds)
        Exit Sub
SaveUserGuildMember_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildMember", Erl)
End Sub

Public Sub SaveUserGuildPedidos(ByVal UserName As String, ByVal Pedidos As String)
        
        On Error GoTo SaveUserGuildPedidos_Err
        
104     Call SaveUserGuildPedidosDatabase(UserName, Pedidos)

        Exit Sub

SaveUserGuildPedidos_Err:
106     Call TraceError(Err.Number, Err.Description, "modGuilds.SaveUserGuildPedidos", Erl)
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
