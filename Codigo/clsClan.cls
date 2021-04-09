VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_GuildId               As Integer
Private p_GuildName             As String
Private p_FounderName           As String
Private p_LeaderName            As String
Private p_Level                 As Byte
Private p_Experience            As Long
Private p_Alignment             As ALINEACION_GUILD
Private p_Description           As String

Private p_OnlineMembers         As Collection   'Array de UserIndexes!

Private Const NEWSLENGTH        As Integer = 1024
Private Const DESCLENGTH        As Integer = 256


Public Property Get GuildName() As String
    On Error GoTo GuildName_Err
        
    GuildName = p_GuildName
        
    Exit Property

GuildName_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.GuildName", Erl)
    Resume Next
        
End Property

'
'ALINEACION Y ANTIFACCION
'

Public Property Get Alineacion() As ALINEACION_GUILD
    On Error GoTo Alineacion_Err
        
    Alineacion = p_Alignment

    Exit Property

Alineacion_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.Alineacion", Erl)
    Resume Next
        
End Property

Public Function CambiarAlineacion(ByVal NuevaAlineacion As ALINEACION_GUILD) As Boolean
    On Error GoTo CambiarAlineacion_Err
        
    p_Alignment = NuevaAlineacion
    Call MakeQuery("UPDATE guilds SET alignment = ? WHERE id = ?;", False, p_Alignment, p_GuildId)
        
    Exit Function

CambiarAlineacion_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.CambiarAlineacion", Erl)
    Resume Next
        
End Function

'
'INICIALIZADORES
'

Private Sub Class_Initialize()
    On Error GoTo Class_Initialize_Err
        
    ' set the connection?
    Exit Sub

Class_Initialize_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.Class_Initialize", Erl)
    Resume Next
        
End Sub

Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Err
        
    Set p_OnlineMembers = Nothing
        
    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.Class_Terminate", Erl)
    Resume Next
        
End Sub

Public Sub InitializeFromRecordset(ByRef QueryData As ADODB.Recordset)
    On Error GoTo Inicializar_Err

    p_GuildId = QueryData!Id
    p_GuildName = QueryData!name
    p_FounderName = QueryData!FounderName
    p_LeaderName = QueryData!LeaderName
    p_Level = QueryData!level
    p_Experience = QueryData!Experience
    p_Alignment = QueryData!Alignment
    p_Description = QueryData!Description

    Set p_OnlineMembers = New Collection
    
    Exit Sub

Inicializar_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.Inicializar", Erl)
    Resume Next
        
End Sub

'
'MEMBRESIAS
'

Public Property Get Fundador() As String
    On Error GoTo Fundador_Err
        
    Fundador = p_FounderName
        
    Exit Property

Fundador_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.Fundador", Erl)
    Resume Next
        
End Property

Public Property Get CantidadDeMiembros() As Integer
    On Error GoTo CantidadDeMiembros_Err
    
    CantidadDeMiembros = 0 'TODO: Hacer el query correspondiente
        
    Exit Property

CantidadDeMiembros_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.CantidadDeMiembros", Erl)
    Resume Next
        
End Property

Public Sub SetLeader(ByRef leader As String)
    On Error GoTo SetLeader_Err
    
    p_LeaderName = leader
        
    Exit Sub

SetLeader_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.SetLeader", Erl)
    Resume Next
        
End Sub

Public Property Get GetLeader() As String
    On Error GoTo GetLeader_Err
        
    GetLeader = p_LeaderName
        
    Exit Property

GetLeader_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.GetLeader", Erl)
    Resume Next
        
End Property

Public Function GetMemberList() As String()
    On Error GoTo GetMemberList_Err
        
    Dim list() As String

    '' FIXME: SELECT u.name FROM guild_members m INNER JOIN user u ON u.id = m.user_id AND m.state = 'active' AND deleted_at IS NULL
    GetMemberList = list

        
    Exit Function

GetMemberList_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.GetMemberList", Erl)
    Resume Next
        
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
    On Error GoTo ConectarMiembro_Err
        
    p_OnlineMembers.Add UserIndex

    With UserList(UserIndex)

        ' No avisa cuando loguea un dios
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then
            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & " se ha conectado."))

        End If

    End With
       
    Exit Sub

ConectarMiembro_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.ConectarMiembro", Erl)
    Resume Next
        
End Sub

Public Sub DesconectarMiembro(ByVal UserIndex As Integer)
    On Error GoTo DesConectarMiembro_Err
    
    Dim i As Integer

    For i = 1 To p_OnlineMembers.Count

        If p_OnlineMembers.Item(i) = UserIndex Then
            p_OnlineMembers.Remove i
            
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

DesConectarMiembro_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.DesconectarMiembro", Erl)
    Resume Next
        
End Sub

Public Sub AceptarNuevoMiembro(ByRef nombre As String)
    On Error GoTo AceptarNuevoMiembro_Err
        

    If PersonajeExiste(nombre) Then
        Call SaveUserGuildIndex(nombre, p_GuildId)
        Call SaveUserGuildAspirant(nombre, 0)

    End If

    Exit Sub

AceptarNuevoMiembro_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.AceptarNuevoMiembro", Erl)
    Resume Next
        
End Sub

Public Sub ExpulsarMiembro(ByRef nombre As String)
    On Error GoTo ExpulsarMiembro_Err
        
    Dim UserIndex As Integer
    Dim EsMiembro As Boolean

    If PersonajeExiste(nombre) Then
        UserIndex = GetUserValue(nombre, "id")
        EsMiembro = GetUserValue(nombre, "guild_index") = p_GuildId
        
        If EsMiembro Then
            Call SaveUserGuildIndex(nombre, 0)
            ' Esto quizas en otro lado
            Call MakeQuery("UPDATE guild_memberships SET state = ?, state_explanation = ? WHERE guild_id = ? AND user_id = ? AND state = ?", _
                True, _
                "expelled", _
                "Ha sido expulsado por el lider", _
                p_GuildId, _
                UserIndex, _
                "active" _
            )
            
        End If

    End If
    
    Exit Sub

ExpulsarMiembro_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.ExpulsarMiembro", Erl)
    Resume Next
        
End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()
        
        On Error GoTo GetAspirantes_Err
        

        Dim list() As String

    
        GetAspirantes = list

        
        Exit Function

GetAspirantes_Err:
116     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetAspirantes", Erl)
118     Resume Next
        
End Function

Public Function CantidadAspirantes() As Integer
    On Error GoTo CantidadAspirantes_Err
    
    CantidadAspirantes = GetDBValue("guild_memberships", "COUNT(1)", "guild_id", p_GuildId)
    
    Exit Function

CantidadAspirantes_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.CantidadAspirantes", Erl)
    Resume Next
        
End Function

Public Function DetallesSolicitudAspirante(ByVal NroAspirante As Integer) As String
        
        On Error GoTo DetallesSolicitudAspirante_Err
        
100     DetallesSolicitudAspirante = ""

        
        Exit Function

DetallesSolicitudAspirante_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.DetallesSolicitudAspirante", Erl)
104     Resume Next
        
End Function

Public Function NumeroDeAspirante(ByRef nombre As String) As Integer
        
        On Error GoTo NumeroDeAspirante_Err
        

        Dim i As Integer
    
100     NumeroDeAspirante = 0
    
102     For i = 1 To MAXASPIRANTES

108     Next i

        
        Exit Function

NumeroDeAspirante_Err:
110     Call RegistrarError(Err.Number, Err.Description, "clsClan.NumeroDeAspirante", Erl)
112     Resume Next
        
End Function

Public Sub NuevoAspirante(ByRef nombre As String, ByRef Peticion As String)
        
        On Error GoTo NuevoAspirante_Err
        

        Dim i     As Integer

        Dim OldQ  As String

        Dim OldQI As Integer

100     OldQ = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")

102     If IsNumeric(OldQ) Then
104         OldQI = CInt(OldQ)
        Else
106         OldQI = 0

        End If

108     For i = 1 To MAXASPIRANTES

110         If GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre") = vbNullString Then
112             Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre", nombre)
114             Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Detalle", IIf(Trim$(Peticion) = vbNullString, "Peticion vacia", Peticion))
116             Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", OldQI + 1)
118             Call SaveUserGuildAspirant(nombre, p_GuildNumber)
                Exit Sub

            End If

120     Next i

        
        Exit Sub

NuevoAspirante_Err:
122     Call RegistrarError(Err.Number, Err.Description, "clsClan.NuevoAspirante", Erl)
124     Resume Next
        
End Sub

Public Sub RetirarAspirante(ByRef nombre As String, ByRef NroAspirante As Integer)
        
        On Error GoTo RetirarAspirante_Err
        

        Dim OldQ    As String

        Dim OldQI   As String

        Dim Pedidos As String

        Dim i       As Integer

100     OldQ = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")

102     If IsNumeric(OldQ) Then
104         OldQI = CInt(OldQ)
        Else
106         OldQI = 1

        End If

108     Call SaveUserGuildAspirant(nombre, 0)
110     Pedidos = GetUserGuildPedidos(nombre)

112     If Not InStr(1, Pedidos, p_GuildName, vbTextCompare) > 0 Then
114         If LenB(Pedidos) <> 0 Then
116             Pedidos = Pedidos & ","

            End If

118         Pedidos = Pedidos & p_GuildName
120         Call SaveUserGuildPedidos(nombre, Pedidos)

        End If
    
122     Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", OldQI - 1)

124     For i = NroAspirante To modGuilds.MAXASPIRANTES - 1
126         Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre", GetVar(SOLICITUDESFILE, "SOLICITUD" & (i + 1), "Nombre"))
128         Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & i, "Detalle", GetVar(SOLICITUDESFILE, "SOLICITUD" & (i + 1), "Detalle"))
130     Next i
    
132     Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & modGuilds.MAXASPIRANTES, "Nombre", vbNullString)
134     Call WriteVar(SOLICITUDESFILE, "SOLICITUD" & modGuilds.MAXASPIRANTES, "Detalle", vbNullString)

        
        Exit Sub

RetirarAspirante_Err:
136     Call RegistrarError(Err.Number, Err.Description, "clsClan.RetirarAspirante", Erl)
138     Resume Next
        
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As String
        
        On Error GoTo GetFechaFundacion_Err
        
100     GetFechaFundacion = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Date")

        
        Exit Function

GetFechaFundacion_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetFechaFundacion", Erl)
104     Resume Next
        
End Function


Public Function GetURL() As String
        
        On Error GoTo GetURL_Err
        
100     GetURL = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL")

        
        Exit Function

GetURL_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetURL", Erl)
104     Resume Next
        
End Function

Public Sub SetNivelDeClan(ByRef nivel As Byte)
        
        On Error GoTo SetNivelDeClan_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "NivelDeClan", nivel)

        
        Exit Sub

SetNivelDeClan_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.SetNivelDeClan", Erl)
104     Resume Next
        
End Sub

Public Function GetNivelDeClan() As Byte
        
        On Error GoTo GetNivelDeClan_Err
        
100     GetNivelDeClan = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "NivelDeClan")

        
        Exit Function

GetNivelDeClan_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetNivelDeClan", Erl)
104     Resume Next
        
End Function

Public Sub SetExpActual(ByRef Exp As Integer)
        
        On Error GoTo SetExpActual_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpActual", Exp)

        
        Exit Sub

SetExpActual_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.SetExpActual", Erl)
104     Resume Next
        
End Sub

Public Function GetExpActual() As Integer
        
        On Error GoTo GetExpActual_Err
        
100     GetExpActual = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpActual")

        
        Exit Function

GetExpActual_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetExpActual", Erl)
104     Resume Next
        
End Function

Public Sub SetExpNecesaria(ByRef Exp As Integer)
        
        On Error GoTo SetExpNecesaria_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpNecesaria", Exp)

        
        Exit Sub

SetExpNecesaria_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.SetExpNecesaria", Erl)
104     Resume Next
        
End Sub

Public Function GetExpNecesaria() As Integer
        
        On Error GoTo GetExpNecesaria_Err
        
100     GetExpNecesaria = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpNecesaria")

        
        Exit Function

GetExpNecesaria_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetExpNecesaria", Erl)
104     Resume Next
        
End Function

Public Sub SetGuildNews(ByRef News As String)
        
        On Error GoTo SetGuildNews_Err
        
100     Call ReplaceInvalidChars(News)
    
102     News = Left$(News, NEWSLENGTH)
    
104     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews", News)

        
        Exit Sub

SetGuildNews_Err:
106     Call RegistrarError(Err.Number, Err.Description, "clsClan.SetGuildNews", Erl)
108     Resume Next
        
End Sub

Public Function GetGuildNews() As String
        
        On Error GoTo GetGuildNews_Err
        
100     GetGuildNews = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews")

        
        Exit Function

GetGuildNews_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetGuildNews", Erl)
104     Resume Next
        
End Function

Public Sub SetDesc(ByRef Desc As String)
        
        On Error GoTo SetDesc_Err
        
100     Call ReplaceInvalidChars(Desc)
102     Desc = Left$(Desc, DESCLENGTH)
    
104     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc", Desc)

        
        Exit Sub

SetDesc_Err:
106     Call RegistrarError(Err.Number, Err.Description, "clsClan.SetDesc", Erl)
108     Resume Next
        
End Sub

Public Function GetDesc() As String
        
        On Error GoTo GetDesc_Err
        
100     GetDesc = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc")

        
        Exit Function

GetDesc_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.GetDesc", Erl)
104     Resume Next
        
End Function

'
'
'ELECCIONES
'
'

Public Function EleccionesAbiertas() As Boolean
        
        On Error GoTo EleccionesAbiertas_Err
        

        Dim ee As String

100     ee = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas")
102     EleccionesAbiertas = (ee = "1")     'cualquier otra cosa da falso

        
        Exit Function

EleccionesAbiertas_Err:
104     Call RegistrarError(Err.Number, Err.Description, "clsClan.EleccionesAbiertas", Erl)
106     Resume Next
        
End Function

Public Sub AbrirElecciones()
        
        On Error GoTo AbrirElecciones_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "1")
102     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", DateAdd("d", 1, Now))
104     Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", "0")

        
        Exit Sub

AbrirElecciones_Err:
106     Call RegistrarError(Err.Number, Err.Description, "clsClan.AbrirElecciones", Erl)
108     Resume Next
        
End Sub

Private Sub CerrarElecciones()  'solo pueden cerrarse mediante recuento de votos
        
        On Error GoTo CerrarElecciones_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "0")
102     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", vbNullString)
104     Call Kill(VOTACIONESFILE)   'borramos toda la evidencia ;-)

        
        Exit Sub

CerrarElecciones_Err:
106     Call RegistrarError(Err.Number, Err.Description, "clsClan.CerrarElecciones", Erl)
108     Resume Next
        
End Sub

Public Sub ContabilizarVoto(ByRef Votante As String, ByRef Votado As String)
        
        On Error GoTo ContabilizarVoto_Err
        

        Dim q     As Integer

        Dim Temps As String

100     Temps = GetVar(VOTACIONESFILE, "INIT", "NumVotos")
102     q = IIf(IsNumeric(Temps), CInt(Temps), 0)
104     Call WriteVar(VOTACIONESFILE, "VOTOS", Votante, Votado)
106     Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", CStr(q + 1))

        
        Exit Sub

ContabilizarVoto_Err:
108     Call RegistrarError(Err.Number, Err.Description, "clsClan.ContabilizarVoto", Erl)
110     Resume Next
        
End Sub

Public Function YaVoto(ByRef Votante) As Boolean
        
        On Error GoTo YaVoto_Err
        
100     YaVoto = ((LenB(Trim$(GetVar(VOTACIONESFILE, "VOTOS", Votante)))) <> 0)

        
        Exit Function

YaVoto_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.YaVoto", Erl)
104     Resume Next
        
End Function

Private Function ContarVotos(ByRef CantGanadores As Integer) As String

        Dim q     As Integer

        Dim i     As Integer

        Dim Temps As String

        Dim tempV As String

        Dim d     As diccionario

        On Error GoTo errh

100     ContarVotos = vbNullString
102     CantGanadores = 0
104     Temps = GetVar(MEMBERSFILE, "INIT", "NroMembers")
106     q = IIf(IsNumeric(Temps), CInt(Temps), 0)

108     If q > 0 Then
            'el diccionario tiene clave el elegido y valor la #votos
110         Set d = New diccionario
        
112         For i = 1 To q
                'miembro del clan
114             Temps = GetVar(MEMBERSFILE, "MEMBERS", "Member" & i)
            
                'a quienvoto
116             tempV = GetVar(VOTACIONESFILE, "VOTOS", Temps)
            
                'si voto a alguien contabilizamos el voto
118             If LenB(tempV) <> 0 Then
120                 If Not IsNull(d.At(tempV)) Then  'cuantos votos tiene?
122                     Call d.AtPut(tempV, CInt(d.At(tempV)) + 1)
                    Else
124                     Call d.AtPut(tempV, 1)

                    End If

                End If

126         Next i
    
            'quien quedo con mas votos, y cuantos tuvieron esos votos?
128         ContarVotos = d.MayorValor(CantGanadores)
    
130         Set d = Nothing

        End If
    
        Exit Function
errh:
132     LogError ("clsClan.Contarvotos: " & Err.Description)

134     If Not d Is Nothing Then Set d = Nothing
136     ContarVotos = vbNullString

End Function

Public Function RevisarElecciones() As Boolean
        
        On Error GoTo RevisarElecciones_Err
        

        Dim FechaSufragio As Date

        Dim Temps         As String

        Dim Ganador       As String

        Dim CantGanadores As Integer

        Dim list()        As String

        Dim i             As Long

100     RevisarElecciones = False
102     Temps = Trim$(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan"))
    
104     If LenB(Temps) = 0 Then Exit Function

106     If IsDate(Temps) Then
108         FechaSufragio = CDate(Temps)

110         If FechaSufragio < Now Then     'toca!
112             Ganador = ContarVotos(CantGanadores)

114             If CantGanadores > 1 Then
                    'empate en la votacion
116                 Call SetGuildNews("*Empate en la votacion. " & Ganador & " con " & CantGanadores & " votos ganaron las elecciones del clan.")
118             ElseIf CantGanadores = 1 Then
120                 list = Me.GetMemberList()
                
122                 For i = 0 To UBound(list())

124                     If Ganador = list(i) Then Exit For
126                 Next i
                
128                 If i <= UBound(list()) Then
130                     Call SetGuildNews("*" & Ganador & " gano la eleccion del clan*")
132                     Call Me.SetLeader(Ganador)
134                     RevisarElecciones = True
                    Else
136                     Call SetGuildNews("*" & Ganador & " gano la eleccion del clan pero abandono las filas por lo que la votacion queda desierta*")

                    End If

                Else
138                 Call SetGuildNews("*El periodo de votacion se cerro sin votos*")

                End If
            
140             Call CerrarElecciones
            
            End If

        Else
142         Call LogError("clsClan.RevisarElecciones: tempS is not Date")

        End If

        
        Exit Function

RevisarElecciones_Err:
144     Call RegistrarError(Err.Number, Err.Description, "clsClan.RevisarElecciones", Erl)
146     Resume Next
        
End Function

'/VOTACIONES

'
'RELACIONES
'


Public Function m_Iterador_ProximoUserIndex() As Integer
        
        On Error GoTo m_Iterador_ProximoUserIndex_Err
        
    
100     If p_IteradorOnlineMembers < p_OnlineMembers.Count Then
102         p_IteradorOnlineMembers = p_IteradorOnlineMembers + 1
104         m_Iterador_ProximoUserIndex = p_OnlineMembers.Item(p_IteradorOnlineMembers)
        Else
106         p_IteradorOnlineMembers = 0
108         m_Iterador_ProximoUserIndex = 0

        End If

        
        Exit Function

m_Iterador_ProximoUserIndex_Err:
110     Call RegistrarError(Err.Number, Err.Description, "clsClan.m_Iterador_ProximoUserIndex", Erl)
112     Resume Next
        
End Function

Public Function Iterador_ProximoGM() As Integer
        
        On Error GoTo Iterador_ProximoGM_Err
        
    
100     If p_IteradorOnlineGMs < p_GMsOnline.Count Then
102         p_IteradorOnlineGMs = p_IteradorOnlineGMs + 1
104         Iterador_ProximoGM = p_GMsOnline.Item(p_IteradorOnlineGMs)
        Else
106         p_IteradorOnlineGMs = 0
108         Iterador_ProximoGM = 0

        End If

        
        Exit Function

Iterador_ProximoGM_Err:
110     Call RegistrarError(Err.Number, Err.Description, "clsClan.Iterador_ProximoGM", Erl)
112     Resume Next
        
End Function

'

'
'ADMINISTRATIVAS
'

Public Sub ConectarGM(ByVal UserIndex As Integer)
        
        On Error GoTo ConectarGM_Err
        
100     p_GMsOnline.Add UserIndex

        
        Exit Sub

ConectarGM_Err:
102     Call RegistrarError(Err.Number, Err.Description, "clsClan.ConectarGM", Erl)
104     Resume Next
        
End Sub

Public Sub DesconectarGM(ByVal UserIndex As Integer)
        
        On Error GoTo DesconectarGM_Err
        

        Dim i As Integer

100     For i = 1 To p_GMsOnline.Count

102         If p_GMsOnline.Item(i) = UserIndex Then
104             p_GMsOnline.Remove (i)

            End If

106     Next i

        
        Exit Sub

DesconectarGM_Err:
108     Call RegistrarError(Err.Number, Err.Description, "clsClan.DesconectarGM", Erl)
110     Resume Next
        
End Sub

'
'VARIAS, EXTRAS Y DEMASES
'

Private Sub ReplaceInvalidChars(ByRef S As String)
    On Error GoTo ReplaceInvalidChars_Err
        

    If InStrB(S, Chr$(13)) <> 0 Then
        S = Replace(S, Chr$(13), vbNullString)

    End If

    If InStrB(S, Chr$(10)) <> 0 Then
        S = Replace(S, Chr$(10), vbNullString)

    End If

    If InStrB(S, "¬") <> 0 Then
        S = Replace(S, "¬", vbNullString)   'morgo usaba esto como "separador"
    End If

        
    Exit Sub

ReplaceInvalidChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsClan.ReplaceInvalidChars", Erl)
    Resume Next
        
End Sub
