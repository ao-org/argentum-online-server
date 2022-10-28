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
'**************************************************************
' clsClan.cls
'
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

''
' clase clan
'
' Es el "ADO" de los clanes. La interfaz entre el disco y
' el juego. Los datos no se guardan en memoria
' para evitar problemas de sincronizacion, y considerando
' que la performance de estas rutinas NO es critica.
' by el oso :p

Private p_GuildName             As String

Private p_Alineacion            As e_ALINEACION_GUILD

Private p_OnlineMembers         As Collection   'Array de UserIndexes!

Private p_GMsOnline             As Collection

Private p_PropuestasDePaz       As Collection

Private p_PropuestasDeAlianza   As Collection

Private p_IteradorRelaciones    As Integer

Private p_IteradorOnlineMembers As Integer

Private p_IteradorPropuesta     As Integer

Private p_IteradorOnlineGMs     As Integer

Private p_GuildNumber           As Integer      'Numero de guild en el mundo

Private p_Relaciones()          As e_RELACIONES_GUILD 'array de relaciones con los otros clanes

Private GUILDINFOFILE           As String

Private GUILDPATH               As String       'aca pq me es mas comodo setearlo y pq en ningun disenio

Private MEMBERSFILE             As String       'decente la capa de arriba se entera donde estan

Private SOLICITUDESFILE         As String       'los datos fisicamente

Private PROPUESTASFILE          As String

Private RELACIONESFILE          As String

Private VOTACIONESFILE          As String

Private Const NEWSLENGTH        As Integer = 1024

Private Const DESCLENGTH        As Integer = 256

Private Const CODEXLENGTH       As Integer = 256

Public Property Get GuildName() As String
        
        On Error GoTo GuildName_Err
        
100     GuildName = p_GuildName

        
        Exit Property

GuildName_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GuildName", Erl)
104
        
End Property

'
'ALINEACION Y ANTIFACCION
'

Public Property Get Alineacion() As e_ALINEACION_GUILD
        
        On Error GoTo Alineacion_Err
        
100     Alineacion = p_Alineacion

        
        Exit Property

Alineacion_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.Alineacion", Erl)
104
        
End Property

Public Property Get PuntosAntifaccion() As Integer
        
        On Error GoTo PuntosAntifaccion_Err
        
100     PuntosAntifaccion = val(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Antifaccion"))

        
        Exit Property

PuntosAntifaccion_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.PuntosAntifaccion", Erl)
104
        
End Property

Public Property Let PuntosAntifaccion(ByVal p As Integer)
        
        On Error GoTo PuntosAntifaccion_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Antifaccion", CStr(p))

        
        Exit Property

PuntosAntifaccion_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.PuntosAntifaccion", Erl)
104
        
End Property

Public Function CambiarAlineacion(ByVal NuevaAlineacion As e_ALINEACION_GUILD) As Boolean
        
        On Error GoTo CambiarAlineacion_Err
        
100     p_Alineacion = NuevaAlineacion
102     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Alineacion", Alineacion2String(p_Alineacion))

        
        Exit Function

CambiarAlineacion_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.CambiarAlineacion", Erl)
106
        
End Function

'
'INICIALIZADORES
'

Private Sub Class_Initialize()
        
        On Error GoTo Class_Initialize_Err
        
100     GUILDPATH = App.Path & "\GUILDS\"
102     GUILDINFOFILE = GUILDPATH & "guildsinfo.inf"

        
        Exit Sub

Class_Initialize_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.Class_Initialize", Erl)
106
        
End Sub

Private Sub Class_Terminate()
        
        On Error GoTo Class_Terminate_Err
        
100     Set p_OnlineMembers = Nothing
102     Set p_GMsOnline = Nothing
104     Set p_PropuestasDePaz = Nothing
106     Set p_PropuestasDeAlianza = Nothing

        
        Exit Sub

Class_Terminate_Err:
108     Call TraceError(Err.Number, Err.Description, "clsClan.Class_Terminate", Erl)
110
        
End Sub

Public Sub Inicializar(ByVal GuildName As String, ByVal GuildNumber As Integer, ByVal Alineacion As e_ALINEACION_GUILD)
        
        On Error GoTo Inicializar_Err
        

        Dim i As Integer

100     p_GuildName = GuildName
102     p_GuildNumber = GuildNumber
104     p_Alineacion = Alineacion
106     Set p_OnlineMembers = New Collection
108     Set p_GMsOnline = New Collection
110     Set p_PropuestasDePaz = New Collection
112     Set p_PropuestasDeAlianza = New Collection
        'ALLIESFILE = GUILDPATH & p_GuildName & "-Allied.all"
        'ENEMIESFILE = GUILDPATH & p_GuildName & "-enemys.ene"
114     RELACIONESFILE = GUILDPATH & p_GuildName & "-relaciones.rel"
116     MEMBERSFILE = GUILDPATH & p_GuildName & "-members.mem"
118     PROPUESTASFILE = GUILDPATH & p_GuildName & "-propositions.pro"
120     SOLICITUDESFILE = GUILDPATH & p_GuildName & "-solicitudes.sol"
122     VOTACIONESFILE = GUILDPATH & p_GuildName & "-votaciones.vot"
124     p_IteradorOnlineMembers = 0
126     p_IteradorPropuesta = 0
128     p_IteradorOnlineGMs = 0
130     p_IteradorRelaciones = 0
132     ReDim Preserve p_Relaciones(1 To CANTIDADDECLANES) As e_RELACIONES_GUILD

134     For i = 1 To CANTIDADDECLANES
136         p_Relaciones(i) = String2Relacion(GetVar(RELACIONESFILE, "RELACIONES", CStr(i)))
138     Next i

140     For i = 1 To CANTIDADDECLANES

142         If Trim$(GetVar(PROPUESTASFILE, CStr(i), "Pendiente")) = "1" Then

144             Select Case String2Relacion(Trim$(GetVar(PROPUESTASFILE, CStr(i), "Tipo")))

                    Case e_RELACIONES_GUILD.ALIADOS
146                     p_PropuestasDeAlianza.Add i

148                 Case e_RELACIONES_GUILD.PAZ
150                     p_PropuestasDePaz.Add i

                End Select

            End If

152     Next i

        
        Exit Sub

Inicializar_Err:
154     Call TraceError(Err.Number, Err.Description, "clsClan.Inicializar", Erl)
156
        
End Sub

''
' esta TIENE QUE LLAMARSE LUEGO DE INICIALIZAR()
'
' @param Fundador Nombre del fundador del clan
'
Public Sub InicializarNuevoClan(ByRef Fundador As String)
        
        On Error GoTo InicializarNuevoClan_Err
        

        Dim OldQ As String  'string pq al comienzo quizas no hay archivo guildinfo.ini y oldq es ""

        Dim NewQ As Integer

        'para que genere los archivos
100     If Not FileExist(MEMBERSFILE, vbNormal) Then
102         Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", "0")

        End If
    
104     If Not FileExist(SOLICITUDESFILE, vbNormal) Then
106         Call WriteVar(SOLICITUDESFILE, "INIT", "CantSolicitudes", "0")

        End If
    
108     OldQ = GetVar(GUILDINFOFILE, "INIT", "nroguilds")

110     If IsNumeric(OldQ) Then
112         NewQ = CInt(Trim$(OldQ)) + 1
        Else
114         NewQ = 1

        End If

116     Call WriteVar(GUILDINFOFILE, "INIT", "NroGuilds", NewQ & vbNewLine)
    
118     Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "GuildName", p_GuildName)
120     Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Founder", Fundador)
122     Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Date", Date)
124     Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Antifaccion", "0")
126     Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "Alineacion", Alineacion2String(p_Alineacion))
127     Call WriteVar(GUILDINFOFILE, "GUILD" & NewQ, "UltimasElecciones", Now)

        
        Exit Sub

InicializarNuevoClan_Err:
128     Call TraceError(Err.Number, Err.Description, "clsClan.InicializarNuevoClan", Erl)
130
        
End Sub

Public Sub ProcesarFundacionDeOtroClan()
        
        On Error GoTo ProcesarFundacionDeOtroClan_Err
        
100     ReDim Preserve p_Relaciones(1 To CANTIDADDECLANES) As e_RELACIONES_GUILD
102     p_Relaciones(CANTIDADDECLANES) = PAZ

        Exit Sub

ProcesarFundacionDeOtroClan_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.ProcesarFundacionDeOtroClan", Erl)
106
        
End Sub

'
'MEMBRESIAS
'

Public Property Get Fundador() As String
        
        On Error GoTo Fundador_Err
        
100     Fundador = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Founder")

        
        Exit Property

Fundador_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.Fundador", Erl)
104
        
End Property

'Public Property Get JugadoresOnline() As String
'Dim i As Integer
'    'leve violacion de capas x aqui, je
'    For i = 1 To p_OnlineMembers.Count
'        JugadoresOnline = UserList(p_OnlineMembers.Item(i)).Name & "," & JugadoresOnline
'    Next i
'End Property

Public Property Get CantidadDeMiembros() As Integer
        
        On Error GoTo CantidadDeMiembros_Err

102     CantidadDeMiembros = val(GetVar(MEMBERSFILE, "INIT", "NroMembers"))

        
        Exit Property

CantidadDeMiembros_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadDeMiembros", Erl)
106
        
End Property

Public Sub SetLeader(ByRef leader As String)
        
        On Error GoTo SetLeader_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Leader", leader)

        
        Exit Sub

SetLeader_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetLeader", Erl)
104
        
End Sub

Public Function GetLeader() As String
        
        On Error GoTo GetLeader_Err
        
100     GetLeader = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Leader")

        
        Exit Function

GetLeader_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetLeader", Erl)
104
        
End Function

Public Function GetMemberList() As String()
        
        On Error GoTo GetMemberList_Err
        

        Dim OldQ   As Integer

        Dim list() As String

        Dim i      As Long
    
100     OldQ = Me.CantidadDeMiembros
    
102     ReDim list(OldQ - 1) As String
    
104     For i = 1 To OldQ
106         list(i - 1) = UCase$(GetVar(MEMBERSFILE, "Members", "Member" & i))
108     Next i
    
110     GetMemberList = list

        
        Exit Function

GetMemberList_Err:
112     Call TraceError(Err.Number, Err.Description, "clsClan.GetMemberList", Erl)
114
        
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 28/05/2010
        '28/05/2010: ZaMa - No avisa cuando loguea un dios o admin.
        '***************************************************
        
        On Error GoTo ConectarMiembro_Err
        
100     p_OnlineMembers.Add UserIndex

102     With UserList(UserIndex)

            ' No avisa cuando loguea un dios
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
106             Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " se ha conectado.", 9))

            End If

        End With

        
        Exit Sub

ConectarMiembro_Err:
108     Call TraceError(Err.Number, Err.Description, "clsClan.ConectarMiembro", Erl)
110
        
End Sub

Public Sub DesConectarMiembro(ByVal UserIndex As Integer)
        
        On Error GoTo DesConectarMiembro_Err
        

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 28/05/2010
        '28/05/2010: ZaMa - No avisa cuando desloguea un dios o admin.
        '***************************************************
        Dim i As Integer

100     For i = 1 To p_OnlineMembers.Count

102         If p_OnlineMembers.Item(i) = UserIndex Then
104             p_OnlineMembers.Remove i
            
106             With UserList(UserIndex)

                    ' No avisa cuando se desconecta un dios
108                 If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
110                     Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & " se ha desconectado.", 8))

                    End If

                End With

                Exit Sub

            End If

112     Next i

        
        Exit Sub

DesConectarMiembro_Err:
114     Call TraceError(Err.Number, Err.Description, "clsClan.DesConectarMiembro", Erl)
116
        
End Sub

Public Sub AceptarNuevoMiembro(ByRef nombre As String)
        
        On Error GoTo AceptarNuevoMiembro_Err
        

        Dim OldQ  As Integer

        Dim OldQs As String

100     If PersonajeExiste(nombre) Then
102         Call SaveUserGuildIndex(nombre, p_GuildNumber)
104         Call SaveUserGuildAspirant(nombre, 0)

106         OldQs = GetVar(MEMBERSFILE, "INIT", "NroMembers")

108         If IsNumeric(OldQs) Then
110             OldQ = CInt(OldQs)
            Else
112             OldQ = 0

            End If

114         Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ + 1)
116         Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ + 1, nombre)

        End If

        
        Exit Sub

AceptarNuevoMiembro_Err:
118     Call TraceError(Err.Number, Err.Description, "clsClan.AceptarNuevoMiembro", Erl)
120
        
End Sub

Public Sub ExpulsarMiembroBug(ByRef nombre As String)
		'NUEVA SUB POR LAS DUDAS PARA NO ROMPER NADA
        On Error GoTo ExpulsarMiembroBug_Err
        Dim OldQ      As Integer
        Dim Temps     As String
        Dim i         As Integer
        Dim EsMiembro As Boolean
        
        OldQ = CInt(GetVar(MEMBERSFILE, "INIT", "NroMembers"))
        i = 1
        nombre = UCase$(nombre)

        While i <= OldQ And UCase$(Trim$(GetVar(MEMBERSFILE, "Members", "Member" & i))) <> nombre
            i = i + 1
        Wend
        EsMiembro = i <= OldQ
        
        If EsMiembro Then
            Call SaveUserGuildIndex(nombre, 0)
            While i < OldQ
                Temps = GetVar(MEMBERSFILE, "Members", "Member" & i + 1)
                Call WriteVar(MEMBERSFILE, "Members", "Member" & i, Temps)
                i = i + 1
            Wend
            Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ, vbNullString)
            Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ - 1)
        End If
        Exit Sub

ExpulsarMiembroBug_Err:
        Call TraceError(Err.Number, Err.Description, "clsClan.ExpulsarMiembroBug", Erl)
End Sub

Public Sub ExpulsarMiembro(ByRef nombre As String)
        
        On Error GoTo ExpulsarMiembro_Err
        

        Dim OldQ      As Integer

        Dim Temps     As String

        Dim i         As Integer

        Dim EsMiembro As Boolean

        Dim MiembroDe As String

100     If PersonajeExiste(nombre) Then
102         OldQ = CInt(GetVar(MEMBERSFILE, "INIT", "NroMembers"))
104         i = 1
106         nombre = UCase$(nombre)

108         While i <= OldQ And UCase$(Trim$(GetVar(MEMBERSFILE, "Members", "Member" & i))) <> nombre

110             i = i + 1
            Wend
112         EsMiembro = i <= OldQ
        
114         If EsMiembro Then
116             Call SaveUserGuildIndex(nombre, 0)

118             While i < OldQ

120                 Temps = GetVar(MEMBERSFILE, "Members", "Member" & i + 1)
122                 Call WriteVar(MEMBERSFILE, "Members", "Member" & i, Temps)
124                 i = i + 1
                Wend
126             Call WriteVar(MEMBERSFILE, "Members", "Member" & OldQ, vbNullString)
                'seteo la cantidad de miembros nueva
128             Call WriteVar(MEMBERSFILE, "INIT", "NroMembers", OldQ - 1)
                'lo echo a el
130             MiembroDe = GetUserGuildMember(nombre)

132             If Not InStr(1, MiembroDe, p_GuildName, vbTextCompare) > 0 Then
134                 If LenB(MiembroDe) <> 0 Then
136                     MiembroDe = MiembroDe & ","

                    End If

138                 MiembroDe = MiembroDe & p_GuildName
140                 Call SaveUserGuildMember(nombre, MiembroDe)

                End If

            End If

        End If

        
        Exit Sub

ExpulsarMiembro_Err:
142     Call TraceError(Err.Number, Err.Description, "clsClan.ExpulsarMiembro", Erl)
144
        
End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()
        
        On Error GoTo GetAspirantes_Err
        

        Dim OldQ   As Integer

        Dim list() As String

        Dim i      As Long
    
100     OldQ = Me.CantidadAspirantes()
    
102     If OldQ > 1 Then
104         ReDim list(OldQ - 1) As String
        Else
106         ReDim list(0) As String

        End If
    
108     For i = 1 To OldQ
110         list(i - 1) = GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre")
112     Next i
    
114     GetAspirantes = list

        
        Exit Function

GetAspirantes_Err:
116     Call TraceError(Err.Number, Err.Description, "clsClan.GetAspirantes", Erl)
118
        
End Function

Public Function CantidadAspirantes() As Integer
        
        On Error GoTo CantidadAspirantes_Err
        

        Dim Temps As String

100     CantidadAspirantes = 0
102     Temps = GetVar(SOLICITUDESFILE, "INIT", "CantSolicitudes")

104     If Not IsNumeric(Temps) Then
            Exit Function

        End If

106     CantidadAspirantes = CInt(Temps)

        
        Exit Function

CantidadAspirantes_Err:
108     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadAspirantes", Erl)
110
        
End Function

Public Function DetallesSolicitudAspirante(ByVal NroAspirante As Integer) As String
        
        On Error GoTo DetallesSolicitudAspirante_Err
        
100     DetallesSolicitudAspirante = GetVar(SOLICITUDESFILE, "SOLICITUD" & NroAspirante, "Detalle")

        
        Exit Function

DetallesSolicitudAspirante_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.DetallesSolicitudAspirante", Erl)
104
        
End Function

Public Function NumeroDeAspirante(ByRef nombre As String) As Integer
        
        On Error GoTo NumeroDeAspirante_Err
        

        Dim i As Integer
    
100     NumeroDeAspirante = 0
    
102     For i = 1 To MAXASPIRANTES

104         If UCase$(Trim$(GetVar(SOLICITUDESFILE, "SOLICITUD" & i, "Nombre"))) = UCase$(nombre) Then
106             NumeroDeAspirante = i
                Exit Function

            End If

108     Next i

        
        Exit Function

NumeroDeAspirante_Err:
110     Call TraceError(Err.Number, Err.Description, "clsClan.NumeroDeAspirante", Erl)
112
        
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
122     Call TraceError(Err.Number, Err.Description, "clsClan.NuevoAspirante", Erl)
124
        
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
136     Call TraceError(Err.Number, Err.Description, "clsClan.RetirarAspirante", Erl)
138
        
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As String
        
        On Error GoTo GetFechaFundacion_Err
        
100     GetFechaFundacion = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Date")

        
        Exit Function

GetFechaFundacion_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetFechaFundacion", Erl)
104
        
End Function

Public Sub SetCodex(ByVal CodexNumber As Integer, ByRef codex As String)
        
        On Error GoTo SetCodex_Err
        
100     Call ReplaceInvalidChars(codex)
102     codex = Left$(codex, CODEXLENGTH)
104     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Codex" & CodexNumber, codex)

        
        Exit Sub

SetCodex_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.SetCodex", Erl)
108
        
End Sub

Public Function GetCodex(ByVal CodexNumber As Integer) As String
        
        On Error GoTo GetCodex_Err
        
100     GetCodex = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Codex" & CodexNumber)

        
        Exit Function

GetCodex_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetCodex", Erl)
104
        
End Function

Public Sub SetURL(ByRef URL As String)
        
        On Error GoTo SetURL_Err
                                                                              
                                                                              
100     Call ReplaceInvalidChars(URL)
        
110     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL", Left$(URL, 40))

        
        Exit Sub

SetURL_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetURL", Erl)
104
        
End Sub

Public Function GetURL() As String
        
        On Error GoTo GetURL_Err
        
100     GetURL = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "URL")

        
        Exit Function

GetURL_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetURL", Erl)
104
        
End Function

Public Sub SetNivelDeClan(ByRef nivel As Byte)
        
        On Error GoTo SetNivelDeClan_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "NivelDeClan", nivel)

        
        Exit Sub

SetNivelDeClan_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetNivelDeClan", Erl)
104
        
End Sub

Public Function GetNivelDeClan() As Byte
        
        On Error GoTo GetNivelDeClan_Err
        
100     GetNivelDeClan = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "NivelDeClan")

        
        Exit Function

GetNivelDeClan_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetNivelDeClan", Erl)
104
        
End Function

Public Sub SetExpActual(ByRef Exp As Integer)
        
        On Error GoTo SetExpActual_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpActual", Exp)

        
        Exit Sub

SetExpActual_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetExpActual", Erl)
104
        
End Sub

Public Function GetExpActual() As Integer
        
        On Error GoTo GetExpActual_Err
        
100     GetExpActual = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpActual")

        
        Exit Function

GetExpActual_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetExpActual", Erl)
104
        
End Function

Public Sub SetExpNecesaria(ByRef Exp As Integer)
        
        On Error GoTo SetExpNecesaria_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpNecesaria", Exp)

        
        Exit Sub

SetExpNecesaria_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetExpNecesaria", Erl)
104
        
End Sub

Public Function GetExpNecesaria() As Integer
        
        On Error GoTo GetExpNecesaria_Err
        
100     GetExpNecesaria = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "ExpNecesaria")

        
        Exit Function

GetExpNecesaria_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetExpNecesaria", Erl)
104
        
End Function

Public Sub SetGuildNews(ByRef News As String)
        
        On Error GoTo SetGuildNews_Err
        
100     Call ReplaceInvalidChars(News)
    
102     News = Left$(News, NEWSLENGTH)
    
104     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews", News)

        
        Exit Sub

SetGuildNews_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.SetGuildNews", Erl)
108
        
End Sub

Public Function GetGuildNews() As String
        
        On Error GoTo GetGuildNews_Err
        
100     GetGuildNews = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "GuildNews")

        
        Exit Function

GetGuildNews_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetGuildNews", Erl)
104
        
End Function

Public Sub SetDesc(ByRef Desc As String)
        
        On Error GoTo SetDesc_Err
        
100     Call ReplaceInvalidChars(Desc)
102     Desc = Left$(Desc, DESCLENGTH)
    
104     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc", Desc)

        
        Exit Sub

SetDesc_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.SetDesc", Erl)
108
        
End Sub

Public Function GetDesc() As String
        
        On Error GoTo GetDesc_Err
        
100     GetDesc = GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "Desc")

        
        Exit Function

GetDesc_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetDesc", Erl)
104
        
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
104     Call TraceError(Err.Number, Err.Description, "clsClan.EleccionesAbiertas", Erl)
106
        
End Function

Public Sub AbrirElecciones()
        
        On Error GoTo AbrirElecciones_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "1")
102     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", DateAdd("d", 1, Now))
103     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "UltimasElecciones", Now)
104     Call WriteVar(VOTACIONESFILE, "INIT", "NumVotos", "0")

        
        Exit Sub

AbrirElecciones_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.AbrirElecciones", Erl)
108
        
End Sub

Private Sub CerrarElecciones()  'solo pueden cerrarse mediante recuento de votos
        
        On Error GoTo CerrarElecciones_Err
        
100     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas", "0")
102     Call WriteVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan", vbNullString)
104     Call Kill(VOTACIONESFILE)   'borramos toda la evidencia ;-)

        
        Exit Sub

CerrarElecciones_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.CerrarElecciones", Erl)
108
        
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
108     Call TraceError(Err.Number, Err.Description, "clsClan.ContabilizarVoto", Erl)
110
        
End Sub

Public Function YaVoto(ByRef Votante) As Boolean
        
        On Error GoTo YaVoto_Err
        
100     YaVoto = ((LenB(Trim$(GetVar(VOTACIONESFILE, "VOTOS", Votante)))) <> 0)

        
        Exit Function

YaVoto_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.YaVoto", Erl)
104
        
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

Public Function RevisarElecciones() As Byte
        
        On Error GoTo RevisarElecciones_Err

        Dim FechaSufragio As Date
        Dim EleccionesAbiertas As Boolean
        Dim Ganador       As String
        Dim CantGanadores As Integer
        Dim list()        As String
        Dim i             As Long
        
100     RevisarElecciones = 0

        EleccionesAbiertas = val(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesAbiertas"))
        
        If EleccionesAbiertas Then 'HarThaoS: Me fijo si hay elecciones abiertas
        
            Dim Temps         As String
            
104         If LenB(Temps) = 0 Then Exit Function

            'Si están abiertas Me fijo cuando finalizan
102         Temps = Trim$(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan"))
106         If IsDate(Temps) Then
108             FechaSufragio = CDate(Temps)
    
110             If FechaSufragio < Now Then     'Si es ahora el sufragio
112                 Ganador = ContarVotos(CantGanadores)
                    If CantGanadores > 1 Then 'Si hay mas de un ganador(empate)
116                     Call SetGuildNews("*Empate en la votacion. " & Ganador & " con " & CantGanadores & " votos ganaron las elecciones del clan.")
                        RevisarElecciones = e_elecciones.HuboEmpate
118                 ElseIf CantGanadores = 1 Then 'Si tenemos un solo ganador
120                     list = Me.GetMemberList()
122                     For i = 0 To UBound(list())
124                         If Ganador = list(i) Then Exit For ' Si el ganador está en el clan
126                     Next i

128                     If i <= UBound(list()) Then 'Aviso que ganó y salgo de aca
130                         Call SetGuildNews("*" & Ganador & " gano la eleccion del clan*")
132                         Call Me.SetLeader(Ganador)
134                         RevisarElecciones = e_elecciones.HayGanador
                        Else 'Se fue del clan
136                         Call SetGuildNews("*" & Ganador & " gano la eleccion del clan pero abandono las filas por lo que la votacion queda desierta.*")
                            RevisarElecciones = e_elecciones.HayGanadorPeroAbandono
                        End If
                    Else 'Si no hubo ganadores quiere decir que no hubo votos. Vuelvo a abrir las elecciones.
138                     Call SetGuildNews("*El periodo de votacion se cerro sin votos*")
                        RevisarElecciones = e_elecciones.NoVotos
                    End If
                    Call CerrarElecciones
                End If
            Else
142             Call LogError("clsClan.RevisarElecciones: tempS is not Date")
            End If
        Else 'Si las elecciones no están abiertas me fijo cuando fue la última vez que se abrieron y si la fecha es anterior a 45 dias la
            Dim UltimasElecciones As String
            
            UltimasElecciones = Trim$(GetVar(GUILDINFOFILE, "GUILD" & p_GuildNumber, "EleccionesFinalizan"))
            
107         If IsDate(UltimasElecciones) Then
                If (UltimasElecciones + DateAdd("d", 45, UltimasElecciones)) >= Now Then
                    Call AbrirElecciones
                    RevisarElecciones = e_elecciones.AbroElecciones
                End If
            Else
                Call LogError("clsClan.RevisarElecciones: UltimasElecciones is not Date")
            End If
                
        End If

        
        Exit Function

RevisarElecciones_Err:
144     Call TraceError(Err.Number, Err.Description, "clsClan.RevisarElecciones", Erl)
146
        
End Function

'/VOTACIONES

'
'RELACIONES
'

Public Property Get CantidadPropuestas(ByVal Tipo As e_RELACIONES_GUILD) As Integer
        
        On Error GoTo CantidadPropuestas_Err
        

100     Select Case Tipo

            Case e_RELACIONES_GUILD.ALIADOS
102             CantidadPropuestas = p_PropuestasDeAlianza.Count

104         Case e_RELACIONES_GUILD.GUERRA
        
106         Case e_RELACIONES_GUILD.PAZ
108             CantidadPropuestas = p_PropuestasDePaz.Count

        End Select

        
        Exit Property

CantidadPropuestas_Err:
110     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadPropuestas", Erl)
112
        
End Property

Public Property Get CantidadEnemys() As Integer
        
        On Error GoTo CantidadEnemys_Err
        

        Dim i As Integer

100     For i = 1 To CANTIDADDECLANES
102         CantidadEnemys = CantidadEnemys + IIf(p_Relaciones(i) = GUERRA, 1, 0)
104     Next i

        
        Exit Property

CantidadEnemys_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadEnemys", Erl)
108
        
End Property

Public Property Get CantidadAllies() As Integer
        
        On Error GoTo CantidadAllies_Err
        

        Dim i As Integer

100     For i = 1 To CANTIDADDECLANES
102         CantidadAllies = CantidadAllies + IIf(p_Relaciones(i) = ALIADOS, 1, 0)
104     Next i

        
        Exit Property

CantidadAllies_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadAllies", Erl)
108
        
End Property

Public Function GetRelacion(ByVal OtroGuild As Integer) As e_RELACIONES_GUILD
        
        On Error GoTo GetRelacion_Err
        
100     GetRelacion = p_Relaciones(OtroGuild)

        
        Exit Function

GetRelacion_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetRelacion", Erl)
104
        
End Function

Public Sub SetRelacion(ByVal GuildIndex As Integer, ByVal Relacion As e_RELACIONES_GUILD)
        
        On Error GoTo SetRelacion_Err
        
100     p_Relaciones(GuildIndex) = Relacion
102     Call WriteVar(RELACIONESFILE, "RELACIONES", GuildIndex, Relacion2String(Relacion))

        
        Exit Sub

SetRelacion_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.SetRelacion", Erl)
106
        
End Sub

Public Sub SetPropuesta(ByVal Tipo As e_RELACIONES_GUILD, ByVal OtroGuild As Integer, ByRef Detalle As String)
        
        On Error GoTo SetPropuesta_Err
        
100     Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle", Detalle)
102     Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Tipo", Relacion2String(Tipo))
104     Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Pendiente", "1")

106     Select Case Tipo

            Case e_RELACIONES_GUILD.ALIADOS
108             p_PropuestasDeAlianza.Add OtroGuild

110         Case e_RELACIONES_GUILD.PAZ
112             p_PropuestasDePaz.Add OtroGuild

        End Select

        
        Exit Sub

SetPropuesta_Err:
114     Call TraceError(Err.Number, Err.Description, "clsClan.SetPropuesta", Erl)
116
        
End Sub

Public Sub AnularPropuestas(ByVal OtroGuild As Integer)
        
        On Error GoTo AnularPropuestas_Err
        

        Dim i As Integer

100     Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle", vbNullString)
102     Call WriteVar(PROPUESTASFILE, CStr(OtroGuild), "Pendiente", "0")

104     For i = 1 To p_PropuestasDePaz.Count

106         If p_PropuestasDePaz.Item(i) = OtroGuild Then p_PropuestasDePaz.Remove (i)
            Exit Sub
108     Next i

110     For i = 1 To p_PropuestasDeAlianza.Count

112         If p_PropuestasDeAlianza.Item(i) = OtroGuild Then p_PropuestasDeAlianza.Remove (i)
            Exit Sub
114     Next i
    
        
        Exit Sub

AnularPropuestas_Err:
116     Call TraceError(Err.Number, Err.Description, "clsClan.AnularPropuestas", Erl)
118
        
End Sub

Public Function GetPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As e_RELACIONES_GUILD) As String
        'trae la solicitd que haya, no valida si es actual o de que tipo es
        
        On Error GoTo GetPropuesta_Err
        
100     GetPropuesta = GetVar(PROPUESTASFILE, CStr(OtroGuild), "Detalle")
102     Tipo = String2Relacion(GetVar(PROPUESTASFILE, CStr(OtroGuild), "Tipo"))

        
        Exit Function

GetPropuesta_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.GetPropuesta", Erl)
106
        
End Function

Public Function HayPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As e_RELACIONES_GUILD) As Boolean
        
        On Error GoTo HayPropuesta_Err
        

        Dim i As Integer

100     HayPropuesta = False

102     Select Case Tipo

            Case e_RELACIONES_GUILD.ALIADOS

104             For i = 1 To p_PropuestasDeAlianza.Count

106                 If p_PropuestasDeAlianza.Item(i) = OtroGuild Then
108                     HayPropuesta = True

                    End If

110             Next i

112         Case e_RELACIONES_GUILD.PAZ

114             For i = 1 To p_PropuestasDePaz.Count

116                 If p_PropuestasDePaz.Item(i) = OtroGuild Then
118                     HayPropuesta = True

                    End If

120             Next i

122         Case e_RELACIONES_GUILD.GUERRA

        End Select

        
        Exit Function

HayPropuesta_Err:
124     Call TraceError(Err.Number, Err.Description, "clsClan.HayPropuesta", Erl)
126
        
End Function

'Public Function GetEnemy(ByVal EnemyIndex As Integer) As String
'    GetEnemy = GetVar(ENEMIESFILE, "ENEMYS", "ENEMY" & EnemyIndex)
'End Function

'Public Function GetAllie(ByVal AllieIndex As Integer) As String
'    GetAllie = GetVar(ALLIESFILE, "ALLIES", "ALLIE" & AllieIndex)
'End Function

'
'ITERADORES
'

Public Function Iterador_ProximaPropuesta(ByVal Tipo As e_RELACIONES_GUILD) As Integer
        
        On Error GoTo Iterador_ProximaPropuesta_Err
        
    
100     Iterador_ProximaPropuesta = 0

102     Select Case Tipo

            Case e_RELACIONES_GUILD.ALIADOS

104             If p_IteradorPropuesta < p_PropuestasDeAlianza.Count Then
106                 p_IteradorPropuesta = p_IteradorPropuesta + 1
108                 Iterador_ProximaPropuesta = p_PropuestasDeAlianza.Item(p_IteradorPropuesta)

                End If
            
110             If p_IteradorPropuesta >= p_PropuestasDeAlianza.Count Then
112                 p_IteradorPropuesta = 0

                End If

114         Case e_RELACIONES_GUILD.PAZ

116             If p_IteradorPropuesta < p_PropuestasDePaz.Count Then
118                 p_IteradorPropuesta = p_IteradorPropuesta + 1
120                 Iterador_ProximaPropuesta = p_PropuestasDePaz.Item(p_IteradorPropuesta)

                End If
            
122             If p_IteradorPropuesta >= p_PropuestasDePaz.Count Then
124                 p_IteradorPropuesta = 0

                End If

        End Select
    
        
        Exit Function

Iterador_ProximaPropuesta_Err:
126     Call TraceError(Err.Number, Err.Description, "clsClan.Iterador_ProximaPropuesta", Erl)
128
        
End Function

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
110     Call TraceError(Err.Number, Err.Description, "clsClan.m_Iterador_ProximoUserIndex", Erl)
112
        
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
110     Call TraceError(Err.Number, Err.Description, "clsClan.Iterador_ProximoGM", Erl)
112
        
End Function

Public Function Iterador_ProximaRelacion(ByVal R As e_RELACIONES_GUILD) As Integer
        
        On Error GoTo Iterador_ProximaRelacion_Err
        
    
100     While p_IteradorRelaciones < UBound(p_Relaciones)
        
102         p_IteradorRelaciones = p_IteradorRelaciones + 1

104         If p_Relaciones(p_IteradorRelaciones) = R Then
106             Iterador_ProximaRelacion = p_IteradorRelaciones
                Exit Function

            End If

        Wend

108     If p_IteradorRelaciones >= UBound(p_Relaciones) Then
110         p_IteradorRelaciones = 0

        End If

        
        Exit Function

Iterador_ProximaRelacion_Err:
112     Call TraceError(Err.Number, Err.Description, "clsClan.Iterador_ProximaRelacion", Erl)
114
        
End Function

'
'
'

'
'ADMINISTRATIVAS
'

Public Sub ConectarGM(ByVal UserIndex As Integer)
        
        On Error GoTo ConectarGM_Err
        
100     p_GMsOnline.Add UserIndex

        
        Exit Sub

ConectarGM_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.ConectarGM", Erl)
104
        
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
108     Call TraceError(Err.Number, Err.Description, "clsClan.DesconectarGM", Erl)
110
        
End Sub

'
'VARIAS, EXTRAS Y DEMASES
'

Private Sub ReplaceInvalidChars(ByRef S As String)
        
        On Error GoTo ReplaceInvalidChars_Err
        

100     If InStrB(S, Chr$(13)) <> 0 Then
102         S = Replace(S, Chr$(13), vbNullString)

        End If

104     If InStrB(S, Chr$(10)) <> 0 Then
106         S = Replace(S, Chr$(10), vbNullString)

        End If

108     If InStrB(S, "¬") <> 0 Then
110         S = Replace(S, "¬", vbNullString)   'morgo usaba esto como "separador"

        End If

        
        Exit Sub

ReplaceInvalidChars_Err:
112     Call TraceError(Err.Number, Err.Description, "clsClan.ReplaceInvalidChars", Erl)
114
        
End Sub

