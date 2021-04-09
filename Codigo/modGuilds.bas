Attribute VB_Name = "modGuilds"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'cantidad maxima de guilds en el servidor
Private Const MAX_GUILDS          As Integer = 1000

'cantidad actual de clanes en el servidor
Public CANTIDADDECLANES           As Integer

'array global de guilds, se indexa por userlist().guildindex
Private guilds(1 To MAX_GUILDS)   As Guild
' FIXME: Esto deberia ser un Dictionary

'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez
Public Const MAXASPIRANTES        As Byte = 10

'alineaciones permitidas
Public Enum ALINEACION_GUILD

    ALINEACION_NINGUNA = 0
    ALINEACION_CIUDADANA = 1
    ALINEACION_CRIMINAL = 2

End Enum

'numero de .wav del cliente
Public Enum eGuildSounds

    Creation = 44
    NewMember = 43
    DeclareWar = 45

End Enum


Public Type Guild
  Id          As Integer
  Name        As String
  FounderId   As Integer
  FounderName As String
  LeaderId    As Integer
  LeaderName  As String
  Level       As Byte
  Experience  As Long
  Alignment   As ALINEACION_GUILD
  Description As String

  ' Extra info that is not persisted
  MembersOnline As Collection
end


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()

    On Error GoTo LoadGuildsDB_Err

    ' FIXME: Extract this query into the GuildRepository class.
    Call MakeQuery("SELECT g.*, l.name AS LeaderName, f.name AS FounderName FROM guilds g INNER JOIN <++> WHERE deleted_at IS NULL;", False)

    If QueryData Is Nothing Then Exit Sub

    CANTIDADDECLANES = QueryData.RecordCount

    QueryData.MoveFirst

    Dim i As Long
    i = 1

    While Not QueryData.EOF
        'FIXME: Deberiamos hacerlo diccionario en lugar de una lista eterna?
        With guilds(i)
          .Id = QueryData!Id
          .GuildName = QueryData!name
          .FounderId = QueryData!FounderId
          .FounderName = QueryData!FounderName
          .LeaderId = QueryData!LeaderId
          .LeaderName = QueryData!LeaderName
          .Level = QueryData!level
          .Experience = QueryData!Experience
          .Alignment = QueryData!Alignment
          .Description = QueryData!Description ' Quizas podriamos NO cargar esto, porque puede ocupar mucho espacio.
        End With

        i = i + 1
        QueryData.MoveNext
    Wend

    Exit Sub

LoadGuildsDB_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGuilds.LoadGuildsDB", Erl)
    Resume Next

End Sub

' When a user goes online and it belongs to a Guild, we let the rest of the guild know.
Public Sub memberConnected(ByVal UserIndex As Integer, ByVal GuildId As Integer)
  On Error GoTo memberConnected_Err

  guilds(GuildId).MembersOnline.Add(UserIndex)

  With UserList(UserIndex)

    ' No avisa cuando loguea un dios
    If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then
      Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & " se ha conectado."))
    End If

  End With

  Exit Sub

memberConnected_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.memberConnected", Erl)
  Resume Next

End Sub

' When a user goes offline and it belongs to a Guild, we let the rest of the guild know.
Public Sub memberDisconnected(ByVal UserIndex As Integer, ByVal GuildId As Integer)
  On Error GoTo memberDisconnected_Err

  Dim i As Long

  For i = 1 To guilds(GuildId).MembersOnline.Count

    If guilds(GuildId).MembersOnline.Item(i) = UserIndex Then
      guilds(GuildId).MembersOnline.Remove i

      With UserList(UserIndex)
        ' No avisa cuando se desconecta un dios
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then
          Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & " se ha desconectado."))

        End If

      End With

      Exit Sub

    End If
  Next i

  Exit Sub

memberDisconnected_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.memberDisconnected", Erl)
  Resume Next

End Sub





Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
  On Error GoTo m_EcharMiembroDeClan_Err

  'UI echa a Expulsado del clan de Expulsado
  Dim UserIndex As Integer
  Dim GI        As Integer
  Dim Map       As Integer

  m_EcharMiembroDeClan = 0

  UserIndex = NameIndex(Expulsado)

  If UserIndex > 0 Then
      'pj online
      GI = UserList(UserIndex).GuildIndex

      If GI > 0 Then
          If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
              If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
              Call guilds(GI).memberDisconnected(UserIndex)
              Call guilds(GI).ExpulsarMiembro(Expulsado)
              Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
              UserList(UserIndex).GuildIndex = 0

              Map = UserList(UserIndex).Pos.Map

              If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
                  Call WriteConsoleMsg(UserIndex, "Necesitas un clan para permanecer en este mapa.", FontTypeNames.FONTTYPE_INFO)
                  Call WarpUserChar(UserIndex, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.X, MapInfo(Map).Salida.Y, True)
              Else
                  Call RefreshCharStatus(UserIndex)
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

       GI = GetUserGuildIndex(Expulsado)

      If GI > 0 Then
          If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
              If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
              Call guilds(GI).ExpulsarMiembro(Expulsado)
              Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)

              Map = GetMapDatabase(Expulsado)

              If MapInfo(Map).SoloClanes And MapInfo(Map).Salida.Map <> 0 Then
                  Call SetPositionDatabase(Expulsado, MapInfo(Map).Salida.Map, MapInfo(Map).Salida.X, MapInfo(Map).Salida.Y)
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
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EcharMiembroDeClan", Erl)
  Resume Next

End Function

Public Sub changeDescription(ByRef Desc As String, ByVal GuildIndex As Integer)
  On Error GoTo changeDescription_Err

  If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub

  guilds(GuildIndex).Description = Desc
  ' FIXME: save guild
  ' private saveGuild(ByRef guilds(GuildIndex))

  Exit Sub

changeDescription_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.changeDescription", Erl)
  Resume Next

End Sub

Public Sub CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef Desc As String, ByRef GuildName As String, ByVal Alineacion As ALINEACION_GUILD)
  On Error GoTo CrearNuevoClan_Err

  Dim i As Integer
  Dim errString As String

  If Not PuedeFundarUnClan(FundadorIndex, Alineacion, errString) Then
    Call WriteConsoleMsg(UserIndex, errString, FontTypeNames.FONTTYPE_GUILD)
    Exit Sub

  End If

  If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
    Call WriteConsoleMsg(UserIndex, "Nombre de clan inválido.", FontTypeNames.FONTTYPE_GUILD)
    Exit Sub

  End If

  If YaExiste(GuildName) Then
      refError = "Ya existe un clan con ese nombre."
      Exit Function

  End If

  Dim newGuild As Guild

  With newGuild
    .Name        = GuildName
    .FounderId = UserList(FundadorIndex).Id
    .FounderName = UserList(FundadorIndex).name
    .LeaderId = .FounderId
    .LeaderName  = .FounderName
    .Level       = 1
    .Experience  = 0
    .Alignment   = Alineacion
    .Description = Desc
  End With

  ' FIXME: Persist newGuild
  ' Asegurarnos de que tiene un .Id seteado

  UserList(FundadorIndex).GuildIndex = newGuild.Id
  Call SetUserValue(UserList(FundadorIndex).name, "guild_index", newGuild.Id)
  Call RefreshCharStatus(FundadorIndex)

  Exit Sub

CrearNuevoClan_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.CrearNuevoClan", Erl)
  Resume Next

End Sub

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

Private Function m_PuedeSalirDeClan(ByRef nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
        'sale solo si no es fundador del clan.

        On Error GoTo m_PuedeSalirDeClan_Err


100     m_PuedeSalirDeClan = False

102     If GuildIndex = 0 Then Exit Function

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

Private Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
  On Error GoTo PuedeFundarUnClan_Err

  PuedeFundarUnClan = False

  If UserList(UserIndex).GuildIndex > 0 Then
      refError = "Ya perteneces a un clan, no podés fundar otro"
      Exit Function

  End If

  If UserList(UserIndex).Stats.ELV < 25 Or UserList(UserIndex).Stats.UserSkills(eSkill.liderazgo) < 80 Then
      refError = "Para fundar un clan debes ser nivel 25, tener 80 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1)."
      Exit Function
  End If

  If Not (TieneObjetos(407, 1, UserIndex) And TieneObjetos(408, 1, UserIndex)) Then
      refError = "Para fundar un clan debes ser nivel 25, tener 80 en liderazgo y tener en tu inventario las 2 gemas: Gema Azul(1), Gema Naranja(1)."
      Exit Function

  End If

  Select Case Alineacion

      Case ALINEACION_GUILD.ALINEACION_CIUDADANA

          If Status(UserIndex) = 0 Or Status(UserIndex) = 2 Then
              refError = "Para fundar un clan de ciudadanos no debes ser criminal."
              Exit Function

          End If

      Case ALINEACION_GUILD.ALINEACION_CRIMINAL

          If Status(UserIndex) = 1 Or Status(UserIndex) = 3 Then
              refError = "Para fundar un clan de criminales no debes ser ciudadano."
              Exit Function

          End If

  End Select

  PuedeFundarUnClan = True

  Exit Function

PuedeFundarUnClan_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.PuedeFundarUnClan", Erl)
  Resume Next

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

126             Case ALINEACION_GUILD.ALINEACION_CRIMINAL

128                 m_EstadoPermiteEntrarChar = Promedio = 0 Or Promedio = 2


            End Select

        End If


        Exit Function

m_EstadoPermiteEntrarChar_Err:
134     Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrarChar", Erl)
136     Resume Next

End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    On Error GoTo m_EstadoPermiteEntrar_Err

    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_CIUDADANA

            m_EstadoPermiteEntrar = Status(UserIndex) = 1 Or Status(UserIndex) = 3

        Case ALINEACION_GUILD.ALINEACION_CRIMINAL

            m_EstadoPermiteEntrar = Status(UserIndex) = 0 Or Status(UserIndex) = 2

        Case ALINEACION_GUILD.ALINEACION_NINGUNA

            m_EstadoPermiteEntrar = True

    End Select

    Exit Function

m_EstadoPermiteEntrar_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EstadoPermiteEntrar", Erl)
    Resume Next

End Function

Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    On Error GoTo String2Alineacion_Err

    Select Case S
        Case "Ciudadano"
            String2Alineacion = ALINEACION_CIUDADANA

        Case "Criminal"
            String2Alineacion = ALINEACION_CRIMINAL

        Else
            String2Alineacion = ALINEACION_NINGUNA

    End Select


    Exit Function

String2Alineacion_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGuilds.String2Alineacion", Erl)
    Resume Next

End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    On Error GoTo Alineacion2String_Err

    Select Case Alineacion

        Case ALINEACION_GUILD.ALINEACION_CIUDADANA
            Alineacion2String = "Ciudadano"

        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"

        Else
            String2Alineacion = "Ninguna"

    End Select


    Exit Function

Alineacion2String_Err:
    Call RegistrarError(Err.Number, Err.Description, "modGuilds.Alineacion2String", Erl)
    Resume Next

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

  YaExiste = GetDBValue("guilds", "COUNT(1)", "UCASE(name)", UCASE$(GuildName)) > 0

  Exit Function

YaExiste_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.YaExiste", Erl)
  Resume Next

End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean

        On Error GoTo v_AbrirElecciones_Err


        Dim GuildIndex As Integer

100     v_AbrirElecciones = False

        refError = "Las elecciones estan desactivadas por el momento."
        Exit Function
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
        ' FIXME: Abrir elecciones!!! Faltan las tablas y todo
118     'Call guilds(GuildIndex).AbrirElecciones


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

        refError = "El sistema de votos esta desabilitado por el momento."
        Exit Function

104     If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
106         refError = "Tu no perteneces a ningún clan"
            Exit Function

        End If

        Exit Function

v_UsuarioVota_Err:
132     Call RegistrarError(Err.Number, Err.Description, "modGuilds.v_UsuarioVota", Erl)
134     Resume Next

End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
  On Error GoTo GuildIndex_Err

  GuildIndex = GetDBValue("guilds", "id", "name", GuildName)

  Exit Function

GuildIndex_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildIndex", Erl)
  Resume Next

End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
  On Error GoTo m_ListaDeMiembrosOnline_Err

  Dim i As Long
  Dim currentUserIsGod As Boolean
  Dim includeUser As Boolean
  Dim tmpUserIndex As Integer

  currentUserIsGod = (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)

  If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
    For i = 1 To guilds(GuildIndex).MembersOnline.Count
      tmpUserIndex = guilds(GuildIndex).MembersOnline.Item(i)

      includeUser = (UserList(tmpUserIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0
      'No mostramos dioses y admins
      If tmpUserIndex <> UserIndex And (includeUser Or currentUserIsGod) Then
        m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(tmpUserIndex).name & ","
      End If

    Next i
  End If

  If Len(m_ListaDeMiembrosOnline) > 0 Then
      m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)

  End If

  Exit Function

m_ListaDeMiembrosOnline_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_ListaDeMiembrosOnline", Erl)
  Resume Next

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
108             tStr(i - 1) = guilds(i).Name & "-" & guilds(i).Alineacion
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

  Dim GuildId As Integer

  GuildId = GetDBValue("guilds", "id", "UCASE(name)", UCASE$(GuildName)

  If GuildId > 0 Then

    With guilds(GI)
      Call Protocol.WriteGuildDetails(UserIndex, .Name, .FounderName, .GetFechaFundacion, .GetLeader, .CantidadDeMiembros, Alineacion2String(.Alineacion), .Description, .Level, .Experience, 1)

    End With
  End If

  Exit Sub

SendGuildDetails_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.SendGuildDetails", Erl)
  Resume Next

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

104     With UserList(UserIndex)
106         GI = .GuildIndex

108         guildList = PrepareGuildsList()

110         If GI <= 0 Or GI > CANTIDADDECLANES Then
                'Send the guild list instead
112             Call Protocol.WriteGuildList(UserIndex, guildList)
                Exit Sub

            End If

114         If Not m_EsGuildLeader(.name, GI) Then
                'Send the guild list instead
116             Call modGuilds.SendGuildNews(UserIndex, guildList)
                '            Call WriteGuildMemberInfo(UserIndex, guildList, MemberList)
                ' Call Protocol.WriteGuildList(UserIndex, guildList)
                Exit Sub

            End If

118         MemberList = guilds(GI).GetMemberList()
120         aspirantsList = guilds(GI).GetAspirantes()

122         Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList, guilds(GI).GetNivelDeClan, guilds(GI).GetExpActual, guilds(GI).GetExpNecesaria)

        End With


        Exit Sub

SendGuildLeaderInfo_Err:
124     Call RegistrarError(Err.Number, Err.Description, "modGuilds.SendGuildLeaderInfo", Erl)
126     Resume Next

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

144     Call SendCharacterInfoDatabase(UserIndex, Personaje)

        Exit Sub
Error:

146     If Not PersonajeExiste(Personaje) Then
148         Call LogError("El usuario " & UserList(UserIndex).name & " (" & UserIndex & " ) ha pedido los detalles del personaje " & Personaje & " que no se encuentra.")
        Else
150         Call LogError("[" & Err.Number & "] " & Err.Description & " En la rutina SendDetallesPersonaje, por el usuario " & UserList(UserIndex).name & " (" & UserIndex & " ), pidiendo informacion sobre el personaje " & Personaje)

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
118         refError = "Tu no puedes entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
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

Public Sub approveMembershipRequest(ByVal ApproverIndex As Integer, ByRef NewMemberName As String)
  On Error GoTo approveMembershipRequest_Err

  Dim GuildId      As Integer
  Dim tGI          As Integer
  Dim NroAspirante As Integer
  Dim AspiranteUI  As Integer
  Dim guild As clsGuild

  GuildId = UserList(ApproverIndex).GuildIndex

  If GuildId <= 0 Or GuildId > CANTIDADDECLANES Then
      Call WriteConsoleMsg(ApproverIndex, "No perteneces a ningún clan.", FontTypeNames.FONTTYPE_GUILD)
      Exit Sub
  End If

  If guilds(GuildId).LeaderId <> UserList(ApproverIndex).Id Then
      Call WriteConsoleMsg(ApproverIndex, "Sólo el líder del clan puede aprobar solicitudes.", FontTypeNames.FONTTYPE_GUILD)
      Exit Sub
  End If

  NewMemberId = GetUserValue(Aspirante, "id")

  If NewMemberId > 0 Then
    If m_EstadoPermiteEntrarChar(Aspirante, GuildId) Then
      tGI = GetUserGuildIndex(Aspirante)

      If tGI <> 0 Then
        Call WriteConsoleMsg(ApproverIndex, Aspirante & " ya es parte de otro clan.", FontTypeNames.FONTTYPE_GUILD)

        Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
        Exit Function
      End If

    Else
      refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
      Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
      Exit Function

    End If
  End If

  'FIXME: Configurar Query para cantida de miembros
  If 1 > MiembrosPermite(GuildId) Then
      Call WriteConsoleMsg(ApproverIndex, "La capacidad del clan esta completa.", FontTypeNames.FONTTYPE_GUILD)
      Exit Sub
  End If

  'el pj es aspirante al clan y puede entrar
  'SQL: UPDATE guild_memberships SET
  Call MakeQuery("UPDATE guild_memberships SET state = ?, state_explanation = ? WHERE guild_id = ? AND user_id = ? AND state = ?", True, _
    "approved", _
    "Ha sido aceptada por el líder del clan.", _
    GuildId, _
    NewMemberId, _
    "pending" _
  )
  Call SetUserValue(Aspirante, "guild_index", GuildId)

  AspiranteUI = NameIndex(Aspirante)
  ' If player is online, update tag
  If AspiranteUI > 0 Then
      Call RefreshCharStatus(AspiranteUI)
      Call memberConnected(AspiranteUI, GuildId)
  End If

  Call SendData(SendTarget.ToGuildMembers, GuildId, PrepareMessageConsoleMsg("[" & Aspirante & "] ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
  Call SendData(SendTarget.ToGuildMembers, GuildId, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

  Exit Sub

approveMembershipRequest_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.approveMembershipRequest", Erl)
  Resume Next

End Function

'FIXME: Borrar esta funcion al pedo!
Public Function GuildLeader(ByVal GuildIndex As Integer) As String

        On Error GoTo GuildLeader_Err


100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function
102     GuildLeader = guilds(GuildIndex).LeaderName


        Exit Function

GuildLeader_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildLeader", Erl)
106     Resume Next

End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String

        On Error GoTo GuildAlignment_Err


100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function

102     GuildAlignment = Alineacion2String(guilds(GuildIndex).Alignment)


        Exit Function

GuildAlignment_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.GuildAlignment", Erl)
106     Resume Next

End Function

Public Function NivelDeClan(ByVal GuildIndex As Integer) As Byte

        On Error GoTo NivelDeClan_Err


100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function

102     NivelDeClan = guilds(GuildIndex).Level


        Exit Function

NivelDeClan_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.NivelDeClan", Erl)
106     Resume Next

End Function

Public Function Alineacion(ByVal GuildIndex As Integer) As ALINEACION_GUILD

        On Error GoTo Alineacion_Err


100     If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then Exit Function

102     Alineacion = guilds(GuildIndex).Alignment


        Exit Function

Alineacion_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modGuilds.Alineacion", Erl)
106     Resume Next

End Function

Public Sub CheckClanExp(ByVal UserIndex As Integer, ByVal ExpDar As Integer)

        On Error GoTo CheckClanExp_Err

        Dim ExpActual    As Integer
        Dim ExpNecesaria As Integer
        Dim GI           As Integer
        Dim nivel        As Byte

100     GI = UserList(UserIndex).GuildIndex
102     ExpActual = guilds(GI).Experience
        nivel = guilds(GI).Level
104     ExpNecesaria = ExperienciaNecesaria(nivel)

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

126         ExpActual = ExpActual - ExpNecesaria

128         nivel = nivel + 1

130         Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, PrepareMessageConsoleMsg("Clan> El clan ha subido a nivel " & nivel & ". Nuevos beneficios disponibles.", FontTypeNames.FONTTYPE_GUILD))

        End If

        guilds(GI).Experience = ExpActual
152     guilds(GI).Level = nivel
        'FIXME: Save guilds(GI)

        Exit Sub

CheckClanExp_Err:
154     Call RegistrarError(Err.Number, Err.Description, "modGuilds.CheckClanExp", Erl)
156     Resume Next

End Sub

Public Function ExperienciaNecesaria(ByVal nivel As Integer) As Integer
  Select Case nivel
    Case 1
      ExperienciaNecesaria = 500
    Case 2
      ExperienciaNecesaria = 1000
    Case 3
      ExperienciaNecesaria = 2000
    Case 4
      ExperienciaNecesaria = 3000
    Case Else
      ExperienciaNecesaria = 0
  End Select
End Function


Private Function MiembrosPermite(ByVal GuildId As Integer) As Byte
  On Error GoTo MiembrosPermite_Err

  Select Case guilds(GuildId).Level
    Case 1
      MiembrosPermite = 15

    Case 2
      MiembrosPermite = 20

    Case 3
      MiembrosPermite = 25

    Case 4
      MiembrosPermite = 30

    Case Else
      MiembrosPermite = 30

  End Select

  Exit Function

MiembrosPermite_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.MiembrosPermite", Erl)
  Resume Next

End Function

Public Function GetUserGuildMember(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Returns the guilds the user has been member of
    '***************************************************

    GetUserGuildMember = SanitizeNullValue(GetUserValue(UserName, "guild_member_history"), vbNullString)

End Function

Public Function GetUserGuildAspirant(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Returns the guilds the user has been member of
    '***************************************************

    GetUserGuildAspirant = SanitizeNullValue(GetUserValue(UserName, "guild_aspirant_index"), 0)

End Function

Public Function GetUserGuildRejectionReason(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Returns the reason why the user has not been accepted to the guild
    '***************************************************

    GetUserGuildRejectionReason = SanitizeNullValue(GetUserValue(UserName, "guild_rejected_because"), vbNullString)

End Function

Public Function GetUserGuildPedidos(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Returns the guilds the user asked to be a member of
    '***************************************************

    GetUserGuildPedidos = SanitizeNullValue(GetUserValue(UserName, "guild_requests_history"), vbNullString)

End Function

Public Sub SaveUserGuildRejectionReason(ByVal UserName As String, ByVal Reason As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the rection reason for the user
    '***************************************************

    Call SetUserValue(UserName, "guild_rejected_because", Reason)

End Sub

Public Sub SaveUserGuildIndex(ByVal UserName As String, ByVal GuildIndex As Integer)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guild index
    '***************************************************

    Call SetUserValue(UserName, "guild_index", GuildIndex)

End Sub

Public Sub SaveUserGuildAspirant(ByVal UserName As String, ByVal AspirantIndex As Integer)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guild Aspirant index
    '***************************************************

    Call SetUserValue(UserName, "guild_aspirant_index", AspirantIndex)

End Sub

Public Sub SaveUserGuildMember(ByVal UserName As String, ByVal guilds As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guilds the user has been member of
    '***************************************************

    Call SetUserValue(UserName, "guild_member_history", guilds)

End Sub

Public Sub SaveUserGuildPedidos(ByVal UserName As String, ByVal Pedidos As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the guilds the user has asked to be a member of
    '***************************************************

    Call SetUserValue(UserName, "guild_requests_history", Pedidos)

End Sub

'''''''''''''''''''''''
'' Private functions ''
'''''''''''''''''''''''

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
  On Error GoTo m_EsGuildLeader_Err

  m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).LeaderName)))

  Exit Function

m_EsGuildLeader_Err:
   Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EsGuildLeader", Erl)
   Resume Next

End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
  On Error GoTo m_EsGuildFounder_Err

  m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).FounderName)))

  Exit Function

m_EsGuildFounder_Err:
  Call RegistrarError(Err.Number, Err.Description, "modGuilds.m_EsGuildFounder", Erl)
  Resume Next

End Function
