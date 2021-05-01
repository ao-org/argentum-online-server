Attribute VB_Name = "modGuilds"
'**************************************************************
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

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

Public Enum ALINEACION_GUILD

    ALINEACION_CIUDA = 0
    ALINEACION_CRIMINAL = 1

    ' ALINEACION_NEUTRO = 3
    ' ALINEACION_CIUDA = 4
    ' ALINEACION_ARMADA = 5
    ' ALINEACION_MASTER = 6
End Enum

'alineaciones permitidas

Public Enum SONIDOS_GUILD

    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45

End Enum

'numero de .wav del cliente

Public Enum RELACIONES_GUILD

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

        Dim Alin       As ALINEACION_GUILD
    
100     GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

102     CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
104     If IsNumeric(CantClanes) Then
106         CANTIDADDECLANES = CInt(CantClanes)
        Else
108         CANTIDADDECLANES = 0

        End If
    
110     For i = 1 To CANTIDADDECLANES
112         Set guilds(i) = New clsClan
114         TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
116         Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
118         Call guilds(i).Inicializar(TempStr, i, Alin)
120     Next i
    
        
        Exit Sub

LoadGuildsDB_Err:
122     Call RegistrarError(Err.Number, Err.Description, "modGuilds.LoadGuildsDB", Erl)
124     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_ConectarMiembroAClan", Erl)
112     Resume Next
        
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
        
        On Error GoTo m_DesconectarMiembroDelClan_Err
        

100     If UserList(UserIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
102     Call guilds(GuildIndex).DesConectarMiembro(UserIndex)

        
        Exit Sub

m_DesconectarMiembroDelClan_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_DesconectarMiembroDelClan", Erl)
106     Resume Next
        
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
        
        On Error GoTo m_EsGuildLeader_Err
        
100     m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))

        
        Exit Function

m_EsGuildLeader_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EsGuildLeader", Erl)
104     Resume Next
        
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
        
        On Error GoTo m_EsGuildFounder_Err
        
100     m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))

        
        Exit Function

m_EsGuildFounder_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EsGuildFounder", Erl)
104     Resume Next
        
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
        
        On Error GoTo m_EcharMiembroDeClan_Err
        

        'UI echa a Expulsado del clan de Expulsado
        Dim UserIndex As Integer

        Dim GI        As Integer
        
        Dim Map       As Integer
    
100     m_EcharMiembroDeClan = 0

102     UserIndex = NameIndex(Expulsado)

104     If UserIndex > 0 Then
            'pj online
106         GI = UserList(UserIndex).GuildIndex

108         If GI > 0 Then
110             If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
112                 If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
114                 Call guilds(GI).DesConectarMiembro(UserIndex)
116                 Call guilds(GI).ExpulsarMiembro(Expulsado)
118                 Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
120                 UserList(UserIndex).GuildIndex = 0

122                 Map = UserList(UserIndex).Pos.Map

124                 If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
126                     Call WriteConsoleMsg(UserIndex, "Necesitas un clan para pertenecer en este mapa.", FontTypeNames.FONTTYPE_INFO)
128                     Call WarpUserChar(UserIndex, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.X, MapInfo(Map).Salida.Y, True)
                    Else
130                     Call RefreshCharStatus(UserIndex)
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
138         If Database_Enabled Then
140             GI = GetUserGuildIndexDatabase(Expulsado)
            Else
142             GI = GetGuildIndexFromChar(Expulsado)
            End If

144         If GI > 0 Then
146             If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
148                 If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
150                 Call guilds(GI).ExpulsarMiembro(Expulsado)
152                 Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)

154                 Map = GetMapDatabase(Expulsado)

156                 If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
158                     Call SetPositionDatabase(Expulsado, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.X, MapInfo(Map).Salida.Y)
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
166     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EcharMiembroDeClan", Erl)
168     Resume Next
        
End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
        
        On Error GoTo ActualizarWebSite_Err
        

        Dim GI As Integer

100     GI = UserList(UserIndex).GuildIndex

102     If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
104     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Sub
    
106     Call guilds(GI).SetURL(Web)
    
        
        Exit Sub

ActualizarWebSite_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modGuilds.ActualizarWebSite", Erl)
110     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.ChangeCodexAndDesc", Erl)
108     Resume Next
        
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
        
        On Error GoTo ActualizarNoticias_Err
        

        Dim GI As Integer

100     GI = UserList(UserIndex).GuildIndex
    
102     If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
104     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then Exit Sub
    
106     Call guilds(GI).SetGuildNews(Datos)
        
        
        Exit Sub

ActualizarNoticias_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modGuilds.ActualizarNoticias", Erl)
110     Resume Next
        
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef Desc As String, ByRef GuildName As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
        
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
120         Call guilds(CANTIDADDECLANES).Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
        
            'Damos de alta al clan como nuevo inicializando sus archivos
122         Call guilds(CANTIDADDECLANES).InicializarNuevoClan(UserList(FundadorIndex).name)
        
            'seteamos codex y descripcion
124         Call guilds(CANTIDADDECLANES).SetDesc(Desc)
126         Call guilds(CANTIDADDECLANES).SetGuildNews("¡Bienvenido a " & GuildName & "! Clan creado con alineación : " & Alineacion2String(Alineacion) & ".")
128         Call guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).name)
        
130         Call guilds(CANTIDADDECLANES).SetNivelDeClan(1)
        
132         Call guilds(CANTIDADDECLANES).SetExpActual(0)
        
134         Call guilds(CANTIDADDECLANES).SetExpNecesaria(500)
        
            '"conectamos" al nuevo miembro a la lista de la clase
136         Call guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).name)
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
154     Call RegistrarError(Err.Number, Err.Description, "modGuilds.CrearNuevoClan", Erl)
156     Resume Next
        
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
    
        Dim MemberList() As String
    
104     MemberList = guilds(GuildIndex).GetMemberList()
    
106     ClanNivel = guilds(GuildIndex).GetNivelDeClan
108     ExpAcu = guilds(GuildIndex).GetExpActual
110     ExpNe = guilds(GuildIndex).GetExpNecesaria

112     Call WriteGuildNews(UserIndex, guilds(GuildIndex).GetGuildNews, guildList, MemberList, ClanNivel, ExpAcu, ExpNe)

        
        Exit Sub

SendGuildNews_Err:
114     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SendGuildNews", Erl)
116     Resume Next
        
End Sub

Public Function m_PuedeSalirDeClan(ByRef nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
        'sale solo si no es fundador del clan.
        
        On Error GoTo m_PuedeSalirDeClan_Err
        

100     m_PuedeSalirDeClan = False

102     If GuildIndex = 0 Then Exit Function
    
        'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
104     If QuienLoEchaUI = -1 Then
106         m_PuedeSalirDeClan = True
            Exit Function

        End If

        'cuando UI no puede echar a nombre?
        'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
108     If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.user Then
110         If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).name), GuildIndex) Then
112             If UCase$(UserList(QuienLoEchaUI).name) <> UCase$(nombre) Then      'si no sale voluntariamente...
                    Exit Function

                End If

            End If

        End If

114     m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).Fundador) <> UCase$(nombre)

        
        Exit Function

m_PuedeSalirDeClan_Err:
116     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_PuedeSalirDeClan", Erl)
118     Resume Next
        
End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
        
        On Error GoTo PuedeFundarUnClan_Err
        

100     PuedeFundarUnClan = False

102     If UserList(UserIndex).GuildIndex > 0 Then
104         refError = "Ya perteneces a un clan, no podés fundar otro"
            Exit Function

        End If
    
106     If UserList(UserIndex).Stats.ELV < 25 Or UserList(UserIndex).Stats.UserSkills(eSkill.liderazgo) < 80 Then
108         refError = "Para fundar un clan debes ser nivel 25, tener 80 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1)."
            Exit Function
        End If
    
110     If Not TieneObjetos(407, 1, UserIndex) Then
112          refError = "Para fundar un clan debes ser nivel 25, tener 80 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1)."
            Exit Function

        End If
    
114     If Not TieneObjetos(408, 1, UserIndex) Then
116           refError = "Para fundar un clan debes ser nivel 25, tener 80 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1)."
            Exit Function

        End If
    
118     Select Case Alineacion

            Case ALINEACION_GUILD.ALINEACION_CIUDA

120             If Status(UserIndex) = 0 Or Status(UserIndex) = 2 Then
122                 refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                    Exit Function

                End If

124         Case ALINEACION_GUILD.ALINEACION_CRIMINAL

126             If Status(UserIndex) = 1 Or Status(UserIndex) = 3 Then
128                 refError = "Para fundar un clan de criminales no debes ser ciudadano."
                    Exit Function

                End If

        End Select

130     PuedeFundarUnClan = True
    
        
        Exit Function

PuedeFundarUnClan_Err:
132     Call RegistrarError(Err.Number, Err.Description, "modGuilds.PuedeFundarUnClan", Erl)
134     Resume Next
        
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
116         Promedio = ObtenerCriminal(Personaje)
        
118         Select Case guilds(GuildIndex).Alineacion

                Case ALINEACION_GUILD.ALINEACION_CIUDA

120                 m_EstadoPermiteEntrarChar = Promedio = 1 Or Promedio = 3

122             Case ALINEACION_GUILD.ALINEACION_CRIMINAL

124                 m_EstadoPermiteEntrarChar = Promedio = 0 Or Promedio = 2


            End Select

        End If

        
        Exit Function

m_EstadoPermiteEntrarChar_Err:
126     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrarChar", Erl)
128     Resume Next
        
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
        On Error GoTo m_EstadoPermiteEntrar_Err

100     Select Case guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_CIUDA

102           m_EstadoPermiteEntrar = Status(UserIndex) = 1 Or Status(UserIndex) = 3

104         Case ALINEACION_GUILD.ALINEACION_CRIMINAL

106           m_EstadoPermiteEntrar = Status(UserIndex) = 0 Or Status(UserIndex) = 2

        End Select

        Exit Function

m_EstadoPermiteEntrar_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrar", Erl)
110     Resume Next

End Function

Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
        
        On Error GoTo String2Alineacion_Err

100     Select Case S

            Case "Ciudadano"
102             String2Alineacion = ALINEACION_CIUDA

104         Case "Criminal"
106             String2Alineacion = ALINEACION_CRIMINAL

        End Select

        
        Exit Function

String2Alineacion_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modGuilds.String2Alineacion", Erl)
110     Resume Next
        
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
        
        On Error GoTo Alineacion2String_Err
        

100     Select Case Alineacion

            Case ALINEACION_GUILD.ALINEACION_CIUDA
102             Alineacion2String = "Ciudadano"

104         Case ALINEACION_GUILD.ALINEACION_CRIMINAL
106             Alineacion2String = "Criminal"

        End Select

        
        Exit Function

Alineacion2String_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modGuilds.Alineacion2String", Erl)
110     Resume Next
        
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
        
        On Error GoTo Relacion2String_Err
        

100     Select Case Relacion

            Case RELACIONES_GUILD.ALIADOS
102             Relacion2String = "A"

104         Case RELACIONES_GUILD.GUERRA
106             Relacion2String = "G"

108         Case RELACIONES_GUILD.PAZ
110             Relacion2String = "P"

112         Case Else
114             Relacion2String = "?"

        End Select

        
        Exit Function

Relacion2String_Err:
116     Call RegistrarError(Err.Number, Err.Description, "modGuilds.Relacion2String", Erl)
118     Resume Next
        
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
        
        On Error GoTo String2Relacion_Err
        

100     Select Case UCase$(Trim$(S))

            Case vbNullString, "P"
102             String2Relacion = RELACIONES_GUILD.PAZ

104         Case "G"
106             String2Relacion = RELACIONES_GUILD.GUERRA

108         Case "A"
110             String2Relacion = RELACIONES_GUILD.ALIADOS

112         Case Else
114             String2Relacion = RELACIONES_GUILD.PAZ

        End Select

        
        Exit Function

String2Relacion_Err:
116     Call RegistrarError(Err.Number, Err.Description, "modGuilds.String2Relacion", Erl)
118     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildNameValido", Erl)
116     Resume Next
        
End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
        
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
112     Call RegistrarError(Err.Number, Err.Description, "modGuilds.YaExiste", Erl)
114     Resume Next
        
End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean
        
        On Error GoTo v_AbrirElecciones_Err
        

        Dim GuildIndex As Integer

100     v_AbrirElecciones = False
102     GuildIndex = UserList(UserIndex).GuildIndex
    
104     If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
106         refError = "Tu no perteneces a ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GuildIndex) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
112     If guilds(GuildIndex).EleccionesAbiertas Then
114         refError = "Las elecciones ya están abiertas"
            Exit Function

        End If
    
116     v_AbrirElecciones = True
118     Call guilds(GuildIndex).AbrirElecciones
    
        
        Exit Function

v_AbrirElecciones_Err:
120     Call RegistrarError(Err.Number, Err.Description, "modGuilds.v_AbrirElecciones", Erl)
122     Resume Next
        
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
        
        On Error GoTo v_UsuarioVota_Err
        

        Dim GuildIndex As Integer

        Dim list()     As String

        Dim i          As Long

100     v_UsuarioVota = False
102     GuildIndex = UserList(UserIndex).GuildIndex
    
104     If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
106         refError = "Tu no perteneces a ningún clan"
            Exit Function

        End If

108     If Not guilds(GuildIndex).EleccionesAbiertas Then
110         refError = "No hay elecciones abiertas en tu clan."
            Exit Function

        End If
    
112     list = guilds(GuildIndex).GetMemberList()

114     For i = 0 To UBound(list())

116         If UCase$(Votado) = list(i) Then Exit For
118     Next i
    
120     If i > UBound(list()) Then
122         refError = Votado & " no pertenece al clan"
            Exit Function

        End If
    
124     If guilds(GuildIndex).YaVoto(UserList(UserIndex).name) Then
126         refError = "Ya has votado, no podés cambiar tu voto"
            Exit Function

        End If
    
128     Call guilds(GuildIndex).ContabilizarVoto(UserList(UserIndex).name, Votado)
130     v_UsuarioVota = True

        
        Exit Function

v_UsuarioVota_Err:
132     Call RegistrarError(Err.Number, Err.Description, "modGuilds.v_UsuarioVota", Erl)
134     Resume Next
        
End Function

Public Sub v_RutinaElecciones()

        Dim i As Integer

        On Error GoTo errh

100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))

102     For i = 1 To CANTIDADDECLANES

104         If Not guilds(i) Is Nothing Then
106             If guilds(i).RevisarElecciones Then
108                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo lider de " & guilds(i).GuildName & "!", FontTypeNames.FONTTYPE_SERVER))

                End If

            End If

proximo:
110     Next i

112     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Elecciones revisadas", FontTypeNames.FONTTYPE_SERVER))
        Exit Sub
errh:
114     Call LogError("modGuilds.v_RutinaElecciones():" & Err.Description)

116     Resume proximo

End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
        
        On Error GoTo GetGuildIndexFromChar_Err
        

        'aca si que vamos a violar las capas deliveradamente ya que
        'visual basic no permite declarar metodos de clase
        Dim Temps As String

100     If InStrB(PlayerName, "\") <> 0 Then
102         PlayerName = Replace(PlayerName, "\", vbNullString)

        End If

104     If InStrB(PlayerName, "/") <> 0 Then
106         PlayerName = Replace(PlayerName, "/", vbNullString)

        End If

108     If InStrB(PlayerName, ".") <> 0 Then
110         PlayerName = Replace(PlayerName, ".", vbNullString)

        End If

112     Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")

114     If IsNumeric(Temps) Then
116         GetGuildIndexFromChar = CInt(Temps)
        Else
118         GetGuildIndexFromChar = 0

        End If

        
        Exit Function

GetGuildIndexFromChar_Err:
120     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GetGuildIndexFromChar", Erl)
122     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildIndex", Erl)
114     Resume Next
        
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
        
        On Error GoTo m_ListaDeMiembrosOnline_Err
        

        Dim i As Integer
    
100     If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
102         i = guilds(GuildIndex).m_Iterador_ProximoUserIndex

104         While i > 0

                'No mostramos dioses y admins
106             If i <> UserIndex And ((UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)) Then m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).name & ","
108             i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
            Wend

        End If

110     If Len(m_ListaDeMiembrosOnline) > 0 Then
112         m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)

        End If

        
        Exit Function

m_ListaDeMiembrosOnline_Err:
114     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_ListaDeMiembrosOnline", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "modGuilds.PrepareGuildsList", Erl)
116     Resume Next
        
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
        
        On Error GoTo SendGuildDetails_Err
        

        Dim codex(CANTIDADMAXIMACODEX - 1) As String

        Dim GI                             As Integer

        Dim i                              As Long

100     GI = GuildIndex(GuildName)

102     If GI = 0 Then Exit Sub
    
104     With guilds(GI)
106         Call Protocol.WriteGuildDetails(UserIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, .CantidadDeMiembros, Alineacion2String(.Alineacion), .GetDesc, .GetNivelDeClan, .GetExpActual, .GetExpNecesaria)

        End With

        
        Exit Sub

SendGuildDetails_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SendGuildDetails", Erl)
110     Resume Next
        
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
        
        On Error GoTo SendGuildLeaderInfo_Err
        

        '***************************************************
        'Autor: Mariano Barrou (El Oso)
        'Last Modification: 12/10/06
        'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        '***************************************************
        Dim GI              As Integer

        Dim guildList()     As String

        Dim MemberList()    As String

        Dim aspirantsList() As String
    
100     With UserList(UserIndex)
102         GI = .GuildIndex
        
104         guildList = PrepareGuildsList()
        
106         If GI <= 0 Or GI > CANTIDADDECLANES Then
                'Send the guild list instead
108             Call Protocol.WriteGuildList(UserIndex, guildList)
                Exit Sub

            End If
        
110         If Not m_EsGuildLeader(.name, GI) Then
                'Send the guild list instead
112             Call modGuilds.SendGuildNews(UserIndex, guildList)
                '            Call WriteGuildMemberInfo(UserIndex, guildList, MemberList)
                ' Call Protocol.WriteGuildList(UserIndex, guildList)
                Exit Sub

            End If
        
114         MemberList = guilds(GI).GetMemberList()
116         aspirantsList = guilds(GI).GetAspirantes()
        
118         Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList, guilds(GI).GetNivelDeClan, guilds(GI).GetExpActual, guilds(GI).GetExpNecesaria)

        End With

        
        Exit Sub

SendGuildLeaderInfo_Err:
120     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SendGuildLeaderInfo", Erl)
122     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_Iterador_ProximoUserIndex", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.Iterador_ProximoGM", Erl)
108     Resume Next
        
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
        'itera sobre las propuestas
        
        On Error GoTo r_Iterador_ProximaPropuesta_Err
        
100     r_Iterador_ProximaPropuesta = 0

102     If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
104         r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)

        End If

        
        Exit Function

r_Iterador_ProximaPropuesta_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_Iterador_ProximaPropuesta", Erl)
108     Resume Next
        
End Function

Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
        
        On Error GoTo GMEscuchaClan_Err
        

        Dim GI As Integer

        'listen to no guild at all
100     If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
            'Quit listening to previous guild!!
102         Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
104         guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            Exit Function

        End If
    
        'devuelve el guildindex
106     GI = GuildIndex(GuildName)

108     If GI > 0 Then
110         If UserList(UserIndex).EscucheClan <> 0 Then
112             If UserList(UserIndex).EscucheClan = GI Then
                    'Already listening to them...
114                 Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
116                 GMEscuchaClan = GI
                    Exit Function
                Else
                    'Quit listening to previous guild!!
118                 Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
120                 guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)

                End If

            End If
        
122         Call guilds(GI).ConectarGM(UserIndex)
124         Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
126         GMEscuchaClan = GI
128         UserList(UserIndex).EscucheClan = GI
        Else
130         Call WriteConsoleMsg(UserIndex, "Error, el clan no existe", FontTypeNames.FONTTYPE_GUILD)
132         GMEscuchaClan = 0

        End If
    
        
        Exit Function

GMEscuchaClan_Err:
134     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GMEscuchaClan", Erl)
136     Resume Next
        
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
        'el index lo tengo que tener de cuando me puse a escuchar
        
        On Error GoTo GMDejaDeEscucharClan_Err
        
100     UserList(UserIndex).EscucheClan = 0
102     Call guilds(GuildIndex).DesconectarGM(UserIndex)

        
        Exit Sub

GMDejaDeEscucharClan_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GMDejaDeEscucharClan", Erl)
106     Resume Next
        
End Sub

Public Function r_DeclararGuerra(ByVal UserIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
        
        On Error GoTo r_DeclararGuerra_Err
        

        Dim GI  As Integer

        Dim GIG As Integer

100     r_DeclararGuerra = 0
102     GI = UserList(UserIndex).GuildIndex

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
112     If Trim$(GuildGuerra) = vbNullString Then
114         refError = "No has seleccionado ningún clan"
            Exit Function

        End If

116     GIG = GuildIndex(GuildGuerra)
    
118     If GI = GIG Then
120         refError = "No podés declarar la guerra a tu mismo clan"
            Exit Function

        End If

122     If GIG < 1 Or GIG > CANTIDADDECLANES Then
124         Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
126         refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
            Exit Function

        End If

128     Call guilds(GI).AnularPropuestas(GIG)
130     Call guilds(GIG).AnularPropuestas(GI)
132     Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
134     Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)

136     r_DeclararGuerra = GIG

        
        Exit Function

r_DeclararGuerra_Err:
138     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_DeclararGuerra", Erl)
140     Resume Next
        
End Function

Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
        
        On Error GoTo r_AceptarPropuestaDePaz_Err
        

        'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
        Dim GI  As Integer

        Dim GIG As Integer

100     GI = UserList(UserIndex).GuildIndex

102     If GI <= 0 Or GI > CANTIDADDECLANES Then
104         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
106     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
108         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
110     If Trim$(GuildPaz) = vbNullString Then
112         refError = "No has seleccionado ningún clan"
            Exit Function

        End If

114     GIG = GuildIndex(GuildPaz)
    
116     If GIG < 1 Or GIG > CANTIDADDECLANES Then
118         Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
120         refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
            Exit Function

        End If

122     If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
124         refError = "No estás en guerra con ese clan"
            Exit Function

        End If
    
126     If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
128         refError = "No hay ninguna propuesta de paz para aceptar"
            Exit Function

        End If

130     Call guilds(GI).AnularPropuestas(GIG)
132     Call guilds(GIG).AnularPropuestas(GI)
134     Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
136     Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    
138     r_AceptarPropuestaDePaz = GIG

        
        Exit Function

r_AceptarPropuestaDePaz_Err:
140     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_AceptarPropuestaDePaz", Erl)
142     Resume Next
        
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
        
        On Error GoTo r_RechazarPropuestaDeAlianza_Err
        

        'devuelve el index al clan guildPro
        Dim GI  As Integer

        Dim GIG As Integer

100     r_RechazarPropuestaDeAlianza = 0
102     GI = UserList(UserIndex).GuildIndex
    
104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
112     If Trim$(GuildPro) = vbNullString Then
114         refError = "No has seleccionado ningún clan"
            Exit Function

        End If

116     GIG = GuildIndex(GuildPro)
    
118     If GIG < 1 Or GIG > CANTIDADDECLANES Then
120         Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
122         refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
            Exit Function

        End If
    
124     If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
126         refError = "No hay propuesta de alianza del clan " & GuildPro
            Exit Function

        End If
    
128     Call guilds(GI).AnularPropuestas(GIG)
        'avisamos al otro clan
130     Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
132     r_RechazarPropuestaDeAlianza = GIG

        
        Exit Function

r_RechazarPropuestaDeAlianza_Err:
134     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_RechazarPropuestaDeAlianza", Erl)
136     Resume Next
        
End Function

Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
        
        On Error GoTo r_RechazarPropuestaDePaz_Err
        

        'devuelve el index al clan guildPro
        Dim GI  As Integer

        Dim GIG As Integer

100     r_RechazarPropuestaDePaz = 0
102     GI = UserList(UserIndex).GuildIndex
    
104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
112     If Trim$(GuildPro) = vbNullString Then
114         refError = "No has seleccionado ningún clan"
            Exit Function

        End If

116     GIG = GuildIndex(GuildPro)
    
118     If GIG < 1 Or GIG > CANTIDADDECLANES Then
120         Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
122         refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
            Exit Function

        End If
    
124     If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
126         refError = "No hay propuesta de paz del clan " & GuildPro
            Exit Function

        End If
    
128     Call guilds(GI).AnularPropuestas(GIG)
        'avisamos al otro clan
130     Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
132     r_RechazarPropuestaDePaz = GIG

        
        Exit Function

r_RechazarPropuestaDePaz_Err:
134     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_RechazarPropuestaDePaz", Erl)
136     Resume Next
        
End Function

Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
        
        On Error GoTo r_AceptarPropuestaDeAlianza_Err
        

        'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
        Dim GI  As Integer

        Dim GIG As Integer

100     r_AceptarPropuestaDeAlianza = 0
102     GI = UserList(UserIndex).GuildIndex

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
112     If Trim$(GuildAllie) = vbNullString Then
114         refError = "No has seleccionado ningún clan"
            Exit Function

        End If

116     GIG = GuildIndex(GuildAllie)
    
118     If GIG < 1 Or GIG > CANTIDADDECLANES Then
120         Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
122         refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
            Exit Function

        End If

124     If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
126         refError = "No estás en paz con el clan, solo podés aceptar propuesas de alianzas con alguien que estes en paz."
            Exit Function

        End If
    
128     If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
130         refError = "No hay ninguna propuesta de alianza para aceptar."
            Exit Function

        End If

132     Call guilds(GI).AnularPropuestas(GIG)
134     Call guilds(GIG).AnularPropuestas(GI)
136     Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
138     Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    
140     r_AceptarPropuestaDeAlianza = GIG

        
        Exit Function

r_AceptarPropuestaDeAlianza_Err:
142     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_AceptarPropuestaDeAlianza", Erl)
144     Resume Next
        
End Function

Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
        
        On Error GoTo r_ClanGeneraPropuesta_Err
        

        Dim OtroClanGI As Integer

        Dim GI         As Integer

100     r_ClanGeneraPropuesta = False
    
102     GI = UserList(UserIndex).GuildIndex

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
108     OtroClanGI = GuildIndex(OtroClan)
    
110     If OtroClanGI = GI Then
112         refError = "No podés declarar relaciones con tu propio clan"
            Exit Function

        End If
    
114     If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
116         refError = "El sistema de clanes esta inconsistente, el otro clan no existe!"
            Exit Function

        End If
    
118     If guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
120         refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
            Exit Function

        End If
    
122     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
124         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
        'de acuerdo al tipo procedemos validando las transiciones
126     If Tipo = RELACIONES_GUILD.PAZ Then
128         If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
130             refError = "No estás en guerra con " & OtroClan
                Exit Function

            End If

132     ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
            'por ahora no hay propuestas de guerra
134     ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then

136         If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
138             refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
                Exit Function

            End If

        End If
    
140     Call guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
142     r_ClanGeneraPropuesta = True

        
        Exit Function

r_ClanGeneraPropuesta_Err:
144     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_ClanGeneraPropuesta", Erl)
146     Resume Next
        
End Function

Public Function r_VerPropuesta(ByVal UserIndex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
        
        On Error GoTo r_VerPropuesta_Err
        

        Dim OtroClanGI As Integer

        Dim GI         As Integer
    
100     r_VerPropuesta = vbNullString
102     refError = vbNullString
    
104     GI = UserList(UserIndex).GuildIndex

106     If GI <= 0 Or GI > CANTIDADDECLANES Then
108         refError = "No eres miembro de ningún clan"
            Exit Function

        End If
    
110     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
112         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
114     OtroClanGI = GuildIndex(OtroGuild)
    
116     If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
118         refError = "No existe la propuesta solicitada"
            Exit Function

        End If
    
120     r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
        
        Exit Function

r_VerPropuesta_Err:
122     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_VerPropuesta", Erl)
124     Resume Next
        
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()
        
        On Error GoTo r_ListaDePropuestas_Err
        

        Dim GI            As Integer

        Dim i             As Integer

        Dim proposalCount As Integer

        Dim proposals()   As String
    
100     GI = UserList(UserIndex).GuildIndex
    
102     If GI > 0 And GI <= CANTIDADDECLANES Then

104         With guilds(GI)
106             proposalCount = .CantidadPropuestas(Tipo)
            
                'Resize array to contain all proposals
108             If proposalCount > 0 Then
110                 ReDim proposals(proposalCount - 1) As String
                Else
112                 ReDim proposals(0) As String

                End If
            
                'Store each guild name
114             For i = 0 To proposalCount - 1
116                 proposals(i) = guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
118             Next i

            End With

        End If
    
120     r_ListaDePropuestas = proposals

        
        Exit Function

r_ListaDePropuestas_Err:
122     Call RegistrarError(Err.Number, Err.Description, "modGuilds.r_ListaDePropuestas", Erl)
124     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.Description, "modGuilds.a_RechazarAspiranteChar", Erl)
116     Resume Next
        
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
        
        On Error GoTo a_ObtenerRechazoDeChar_Err
        

100     If InStrB(Aspirante, "\") <> 0 Then
102         Aspirante = Replace(Aspirante, "\", "")

        End If

104     If InStrB(Aspirante, "/") <> 0 Then
106         Aspirante = Replace(Aspirante, "/", "")

        End If

108     If InStrB(Aspirante, ".") <> 0 Then
110         Aspirante = Replace(Aspirante, ".", "")

        End If

112     a_ObtenerRechazoDeChar = GetUserGuildRejectionReason(Aspirante)
114     Call SaveUserGuildRejectionReason(Aspirante, vbNullString)

        
        Exit Function

a_ObtenerRechazoDeChar_Err:
116     Call RegistrarError(Err.Number, Err.Description, "modGuilds.a_ObtenerRechazoDeChar", Erl)
118     Resume Next
        
End Function

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

108     NroAspirante = guilds(GI).NumeroDeAspirante(nombre)

110     If NroAspirante = 0 Then
112         refError = nombre & " no es aspirante a tu clan"
            Exit Function

        End If

114     Call guilds(GI).RetirarAspirante(nombre, NroAspirante)
116     refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
118     a_RechazarAspirante = True

        
        Exit Function

a_RechazarAspirante_Err:
120     Call RegistrarError(Err.Number, Err.Description, "modGuilds.a_RechazarAspirante", Erl)
122     Resume Next
        
End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef nombre As String) As String
        
        On Error GoTo a_DetallesAspirante_Err
        

        Dim GI           As Integer

        Dim NroAspirante As Integer

100     GI = UserList(UserIndex).GuildIndex

102     If GI <= 0 Or GI > CANTIDADDECLANES Then
            Exit Function

        End If
    
104     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
            Exit Function

        End If
    
106     NroAspirante = guilds(GI).NumeroDeAspirante(nombre)

108     If NroAspirante > 0 Then
110         a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)

        End If
    
        
        Exit Function

a_DetallesAspirante_Err:
112     Call RegistrarError(Err.Number, Err.Description, "modGuilds.a_DetallesAspirante", Erl)
114     Resume Next
        
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Dim GI     As Integer

        Dim NroAsp As Integer

        Dim list() As String

        Dim i      As Long
    
        On Error GoTo Error

100     GI = UserList(UserIndex).GuildIndex
    
102     Personaje = UCase$(Personaje)
    
104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         Call Protocol.WriteConsoleMsg(UserIndex, "No perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
110         Call Protocol.WriteConsoleMsg(UserIndex, "No eres el lider de tu clan.", FontTypeNames.FONTTYPE_INFO)
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
    
124     NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    
126     If NroAsp = 0 Then
128         list = guilds(GI).GetMemberList()
        
130         For i = 0 To UBound(list())

132             If Personaje = list(i) Then Exit For
134         Next i
        
136         If i > UBound(list()) Then
138             Call Protocol.WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro del clan.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If

140     Call SendCharacterInfoDatabase(UserIndex, Personaje)

        Exit Sub
Error:

142     If Not PersonajeExiste(Personaje) Then
144         Call LogError("El usuario " & UserList(UserIndex).name & " (" & UserIndex & " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
        Else
146         Call LogError("[" & Err.Number & "] " & Err.Description & " En la rutina SendDetallesPersonaje, por el usuario " & UserList(UserIndex).name & " (" & UserIndex & " ), pidiendo informacion sobre el personaje " & Personaje)

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
            
128         If UserList(UserIndex).name = NuevoGuildAspirantes(i) Then
130             refError = "Ya has enviado una solicitud a este clan."
                Exit Function

            End If
                    
        Next

132     ViejoSolicitado = GetVar(CharPath & UserList(UserIndex).name & ".chr", "GUILD", "ASPIRANTEA")

134     If LenB(ViejoSolicitado) <> 0 Then
            'borramos la vieja solicitud
136         ViejoGuildINdex = CInt(ViejoSolicitado)

138         If ViejoGuildINdex <> 0 Then
140             ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).name)

142             If ViejoNroAspirante > 0 Then
144                 Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).name, ViejoNroAspirante)

                End If

            Else

                'RefError = "Inconsistencia en los clanes, avise a un administrador"
                'Exit Function
            End If

        End If
    
146     Call SendData(SendTarget.ToDiosesYclan, NuevoGuildIndex, PrepareMessageGuildChat("Clan> [" & UserList(UserIndex).name & "] ha enviado solicitud para unirse al clan."))
    
148     Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).name, Solicitud)
150     a_NuevoAspirante = True

        
        Exit Function

a_NuevoAspirante_Err:
152     Call RegistrarError(Err.Number, Err.Description, "modGuilds.a_NuevoAspirante", Erl)
154     Resume Next
        
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
        
        On Error GoTo a_AceptarAspirante_Err
        

        Dim GI           As Integer
        
        Dim tGI          As Integer

        Dim NroAspirante As Integer

        Dim AspiranteUI  As Integer

        'un pj ingresa al clan :D

100     a_AceptarAspirante = False
    
102     GI = UserList(UserIndex).GuildIndex

104     If GI <= 0 Or GI > CANTIDADDECLANES Then
106         refError = "No perteneces a ningún clan"
            Exit Function

        End If
    
108     If Not m_EsGuildLeader(UserList(UserIndex).name, GI) Then
110         refError = "No eres el líder de tu clan"
            Exit Function

        End If
    
112     NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    
114     If NroAspirante = 0 Then
116         refError = "El Pj no es aspirante al clan"
            Exit Function

        End If
    
118     AspiranteUI = NameIndex(Aspirante)

120     If AspiranteUI > 0 Then

            'pj Online
122         If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
124             refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
126             Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
                Exit Function
128         ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
130             refError = Aspirante & " ya es parte de otro clan."
132             Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
                Exit Function

            End If

        Else

134         If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
136             refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
138             Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
                Exit Function
            Else
140             If Database_Enabled Then
142                 tGI = GetUserGuildIndexDatabase(Aspirante)
                Else
144                 tGI = GetGuildIndexFromChar(Aspirante)
                End If
                
146             If tGI <> 0 Then
148                 refError = Aspirante & " ya es parte de otro clan."
150                 Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
                    Exit Function
                End If

            End If

        End If
    
152     If guilds(GI).CantidadDeMiembros + 1 > MiembrosPermite(GI) Then
154         refError = "La capacidad del clan esta completa."
            Exit Function

        End If
    
        'el pj es aspirante al clan y puede entrar
    
156     Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
158     Call guilds(GI).AceptarNuevoMiembro(Aspirante)
    
        ' If player is online, update tag
160     If AspiranteUI > 0 Then
162         Call RefreshCharStatus(AspiranteUI)

        End If
    
164     a_AceptarAspirante = True

        
        Exit Function

a_AceptarAspirante_Err:
166     Call RegistrarError(Err.Number, Err.Description, "modGuilds.a_AceptarAspirante", Erl)
168     Resume Next
        
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
        
        On Error GoTo GuildName_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     GuildName = guilds(GuildIndex).GuildName

        
        Exit Function

GuildName_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildName", Erl)
106     Resume Next
        
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
        
        On Error GoTo GuildLeader_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
102     GuildLeader = guilds(GuildIndex).GetLeader

        
        Exit Function

GuildLeader_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildLeader", Erl)
106     Resume Next
        
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
        
        On Error GoTo GuildAlignment_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)

        
        Exit Function

GuildAlignment_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildAlignment", Erl)
106     Resume Next
        
End Function

Public Function NivelDeClan(ByVal GuildIndex As Integer) As Byte
        
        On Error GoTo NivelDeClan_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     NivelDeClan = guilds(GuildIndex).GetNivelDeClan

        
        Exit Function

NivelDeClan_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.NivelDeClan", Erl)
106     Resume Next
        
End Function

Public Function Alineacion(ByVal GuildIndex As Integer) As Byte
        
        On Error GoTo Alineacion_Err
        

100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
    
102     Alineacion = guilds(GuildIndex).Alineacion

        
        Exit Function

Alineacion_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.Alineacion", Erl)
106     Resume Next
        
End Function

Sub CheckClanExp(ByVal UserIndex As Integer, ByVal ExpDar As Integer)
        
        On Error GoTo CheckClanExp_Err
        

        Dim ExpActual    As Integer

        Dim ExpNecesaria As Integer

        Dim GI           As Integer

        Dim nivel        As Byte

100     GI = UserList(UserIndex).GuildIndex
102     ExpActual = guilds(GI).GetExpActual
104     ExpNecesaria = guilds(GI).GetExpNecesaria
106     nivel = guilds(GI).GetNivelDeClan

108     If nivel >= 5 Then
            Exit Sub

        End If

110     If UserList(UserIndex).ChatCombate = 1 Then
112         Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, PrepareMessageConsoleMsg("Clan> El clan ha ganado " & ExpDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_GUILD))

        End If

114     ExpActual = ExpActual + ExpDar

116     If ExpActual >= ExpNecesaria Then
    
            'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
            'nivel
118         If nivel >= 5 Then
120             ExpActual = 0
122             ExpNecesaria = 0
                Exit Sub

            End If

124         Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, PrepareMessagePlayWave(SND_NIVEL, NO_3D_SOUND, NO_3D_SOUND))
    
            ' UserList(UserIndex).Familiar.Exp = UserList(UserIndex).Familiar.Exp - UserList(UserIndex).Familiar.ELU
    
126         ExpActual = ExpActual - ExpNecesaria
    
128         nivel = nivel + 1
    
130         Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, PrepareMessageConsoleMsg("Clan> El clan ha subido a nivel " & nivel & ". Nuevos beneficios disponibles.", FontTypeNames.FONTTYPE_GUILD))
    
            '  UserList(UserIndex).Familiar.Exp = UserList(UserIndex).Familiar.Exp - UserList(UserIndex).Familiar.ELU
    
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
132         If nivel = 2 Then
134             ExpNecesaria = 1000
136         ElseIf nivel = 3 Then
138             ExpNecesaria = 2000
140         ElseIf nivel = 4 Then
142             ExpNecesaria = 3000
            Else
144             ExpNecesaria = 0
146             ExpActual = 0

            End If
   
            ' guilds(gi).SetExpNecesaria = ExpNecesaria
            'guilds(gi).SetExpActual = ExpActual
            ' guilds(gi).SetNivelDeClan = nivel

        End If

148     guilds(GI).SetExpNecesaria (ExpNecesaria)
150     guilds(GI).SetExpActual (ExpActual)
152     guilds(GI).SetNivelDeClan (nivel)

        
        Exit Sub

CheckClanExp_Err:
154     Call RegistrarError(Err.Number, Err.Description, "modGuilds.CheckClanExp", Erl)
156     Resume Next
        
End Sub

Public Function MiembrosPermite(ByVal GI As Integer) As Byte
        
        On Error GoTo MiembrosPermite_Err
        

        Dim nivel As Byte

100     nivel = guilds(GI).GetNivelDeClan

102     Select Case nivel

            Case 1
104             MiembrosPermite = 15

106         Case 2
108             MiembrosPermite = 20

110         Case 3
112             MiembrosPermite = 25

114         Case 4
116             MiembrosPermite = 30

118         Case Else
120             MiembrosPermite = 30

        End Select

        
        Exit Function

MiembrosPermite_Err:
122     Call RegistrarError(Err.Number, Err.Description, "modGuilds.MiembrosPermite", Erl)
124     Resume Next
        
End Function

Public Function GetUserGuildMember(ByVal UserName As String) As String
        
        On Error GoTo GetUserGuildMember_Err
        

        '***************************************************
        'Author: Juan Andres Dalmasso
        'Returns the guilds the user has been member of
        '***************************************************
100     If Not Database_Enabled Then
102         GetUserGuildMember = GetUserGuildMemberCharfile(UserName)
        Else
104         GetUserGuildMember = GetUserGuildMemberDatabase(UserName)

        End If

        
        Exit Function

GetUserGuildMember_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GetUserGuildMember", Erl)
108     Resume Next
        
End Function

Public Function GetUserGuildAspirant(ByVal UserName As String) As Integer
        
        On Error GoTo GetUserGuildAspirant_Err
        

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 24/09/2018
        'Returns the guilds the user has been member of
        '***************************************************
100     If Not Database_Enabled Then
102         GetUserGuildAspirant = GetUserGuildAspirantCharfile(UserName)
        Else
104         GetUserGuildAspirant = GetUserGuildAspirantDatabase(UserName)

        End If

        
        Exit Function

GetUserGuildAspirant_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GetUserGuildAspirant", Erl)
108     Resume Next
        
End Function

Public Function GetUserGuildRejectionReason(ByVal UserName As String) As String
        
        On Error GoTo GetUserGuildRejectionReason_Err
        

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 24/09/2018
        'Returns the reason why the user has not been accepted to the guild
        '***************************************************
100     If Not Database_Enabled Then
102         GetUserGuildRejectionReason = GetUserGuildRejectionReasonCharfile(UserName)
        Else
104         GetUserGuildRejectionReason = GetUserGuildRejectionReasonDatabase(UserName)

        End If

        
        Exit Function

GetUserGuildRejectionReason_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GetUserGuildRejectionReason", Erl)
108     Resume Next
        
End Function

Public Function GetUserGuildPedidos(ByVal UserName As String) As String
        
        On Error GoTo GetUserGuildPedidos_Err
        

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 24/09/2018
        'Returns the guilds the user asked to be a member of
        '***************************************************
100     If Not Database_Enabled Then
102         GetUserGuildPedidos = GetUserGuildPedidosCharfile(UserName)
        Else
104         GetUserGuildPedidos = GetUserGuildPedidosDatabase(UserName)

        End If

        
        Exit Function

GetUserGuildPedidos_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GetUserGuildPedidos", Erl)
108     Resume Next
        
End Function

Public Sub SaveUserGuildRejectionReason(ByVal UserName As String, ByVal Reason As String)
        
        On Error GoTo SaveUserGuildRejectionReason_Err
        

        '***************************************************
        'Autor: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        'Updates the rection reason for the user
        '***************************************************
100     If Not Database_Enabled Then
102         Call SaveUserGuildRejectionReasonCharfile(UserName, Reason)
        Else
104         Call SaveUserGuildRejectionReasonDatabase(UserName, Reason)

        End If

        
        Exit Sub

SaveUserGuildRejectionReason_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SaveUserGuildRejectionReason", Erl)
108     Resume Next
        
End Sub

Public Sub SaveUserGuildIndex(ByVal UserName As String, ByVal GuildIndex As Integer)
        
        On Error GoTo SaveUserGuildIndex_Err
        

        '***************************************************
        'Autor: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        'Updates the guild index
        '***************************************************
100     If Not Database_Enabled Then
102         Call SaveUserGuildIndexCharfile(UserName, GuildIndex)
        Else
104         Call SaveUserGuildIndexDatabase(UserName, GuildIndex)

        End If

        
        Exit Sub

SaveUserGuildIndex_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SaveUserGuildIndex", Erl)
108     Resume Next
        
End Sub

Public Sub SaveUserGuildAspirant(ByVal UserName As String, ByVal AspirantIndex As Integer)
        
        On Error GoTo SaveUserGuildAspirant_Err
        

        '***************************************************
        'Autor: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        'Updates the guild Aspirant index
        '***************************************************
100     If Not Database_Enabled Then
102         Call SaveUserGuildAspirantCharfile(UserName, AspirantIndex)
        Else
104         Call SaveUserGuildAspirantDatabase(UserName, AspirantIndex)

        End If

        
        Exit Sub

SaveUserGuildAspirant_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SaveUserGuildAspirant", Erl)
108     Resume Next
        
End Sub

Public Sub SaveUserGuildMember(ByVal UserName As String, ByVal guilds As String)
        
        On Error GoTo SaveUserGuildMember_Err
        

        '***************************************************
        'Autor: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        'Updates the guilds the user has been member of
        '***************************************************
100     If Not Database_Enabled Then
102         Call SaveUserGuildMemberCharfile(UserName, guilds)
        Else
104         Call SaveUserGuildMemberDatabase(UserName, guilds)

        End If

        
        Exit Sub

SaveUserGuildMember_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SaveUserGuildMember", Erl)
108     Resume Next
        
End Sub

Public Sub SaveUserGuildPedidos(ByVal UserName As String, ByVal Pedidos As String)
        
        On Error GoTo SaveUserGuildPedidos_Err
        

        '***************************************************
        'Autor: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        'Updates the guilds the user has asked to be a member of
        '***************************************************
100     If Not Database_Enabled Then
102         Call SaveUserGuildPedidosCharfile(UserName, Pedidos)
        Else
104         Call SaveUserGuildPedidosDatabase(UserName, Pedidos)

        End If

        
        Exit Sub

SaveUserGuildPedidos_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SaveUserGuildPedidos", Erl)
108     Resume Next
        
End Sub
