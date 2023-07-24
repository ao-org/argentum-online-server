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
Private p_Leader                As Long
Private p_Founder               As Long
Private p_IteradorOnlineMembers As Integer
Private p_IteradorPropuesta     As Integer
Private p_IteradorOnlineGMs     As Integer
Private p_GuildNumber           As Integer      'Numero de guild en el mundo
Private p_GuildId               As Long
Private Const NEWSLENGTH        As Integer = 1024
Private Const DESCLENGTH        As Integer = 256
Private Const CODEXLENGTH       As Integer = 256
Private MemberList()            As Long
Private MemberCount             As Integer
Private MemberListLoaded        As Boolean

Public Property Get GuildName() As String
        On Error GoTo GuildName_Err
100     GuildName = p_GuildName
        Exit Property
GuildName_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GuildName", Erl)
End Property

Private Sub UpdateGuildValue(ByVal ColumnName As String, ByVal Value As Variant)
    Call Execute("UPDATE guilds SET " & ColumnName & " = ? WHERE id = ?;", Value, p_GuildId)
End Sub
'
'ALINEACION Y ANTIFACCION
'
Public Property Get Alineacion() As e_ALINEACION_GUILD
        On Error GoTo Alineacion_Err
100     Alineacion = p_Alineacion
        Exit Property
Alineacion_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.Alineacion", Erl)
End Property

Public Function CambiarAlineacion(ByVal NuevaAlineacion As e_ALINEACION_GUILD) As Boolean
        On Error GoTo CambiarAlineacion_Err
100     p_Alineacion = NuevaAlineacion
        Call UpdateGuildValue("alignment", p_Alineacion)
        Exit Function
CambiarAlineacion_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.CambiarAlineacion", Erl)
End Function

Private Sub Class_Initialize()
    Set p_OnlineMembers = New Collection
    Set p_GMsOnline = New Collection
    p_IteradorOnlineMembers = 0
    p_IteradorOnlineGMs = 0
    MemberListLoaded = False
End Sub

Private Sub Class_Terminate()
        On Error GoTo Class_Terminate_Err
100     Set p_OnlineMembers = Nothing
102     Set p_GMsOnline = Nothing
        Exit Sub
Class_Terminate_Err:
108     Call TraceError(Err.Number, Err.Description, "clsClan.Class_Terminate", Erl)
End Sub

Public Sub InitFromRecord(RS As Recordset, ByVal GuildNumber As Integer)
On Error GoTo Inicializar_Err
    p_GuildName = RS!guild_name
    p_GuildId = RS!id
    p_Alineacion = RS!alignment
    p_GuildNumber = GuildNumber
    p_Leader = RS!leader_id
    p_Founder = RS!founder_id
    Exit Sub
Inicializar_Err:
154     Call TraceError(Err.Number, Err.Description, "clsClan.Inicializar", Erl)
End Sub


' @param Fundador Nombre del fundador del clan
'
Public Sub InicializarNuevoClan(ByVal GuildName As String, ByVal GuildNumber As Integer, ByVal alignment As e_ALINEACION_GUILD, ByVal founder As Long)
    On Error GoTo InicializarNuevoClan_Err
        p_GuildName = GuildName
        p_Alineacion = alignment
        p_GuildNumber = GuildNumber
        p_Leader = founder
        p_Founder = founder
        Call Query("INSERT INTO guilds (founder_id, guild_name, alignment, leader_id) VALUES (?,?,?,?)", p_Founder, p_GuildName, p_Alineacion, p_Leader)
        ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("SELECT last_insert_rowid()")
        If RS Is Nothing Then
            p_GuildId = 1
        Else
            p_GuildId = val(RS.Fields(0).Value)
        End If
        Exit Sub
InicializarNuevoClan_Err:
128     Call TraceError(Err.Number, Err.Description, "clsClan.InicializarNuevoClan", Erl)
End Sub

Public Sub ProcesarFundacionDeOtroClan()
        On Error GoTo ProcesarFundacionDeOtroClan_Err
        Exit Sub
ProcesarFundacionDeOtroClan_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.ProcesarFundacionDeOtroClan", Erl)
End Sub

'
'MEMBRESIAS
'
Public Property Get Fundador() As Long
    On Error GoTo Fundador_Err
100     Fundador = p_Founder
        Exit Property
Fundador_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.Fundador", Erl)
End Property

Public Property Get CantidadDeMiembros() As Integer
        On Error GoTo CantidadDeMiembros_Err
        Call ReloadMemberList
102     CantidadDeMiembros = MemberCount
        Exit Property
CantidadDeMiembros_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadDeMiembros", Erl)
End Property

Public Sub SetLeader(ByRef Leader As Long)
    On Error GoTo SetLeader_Err
        Call UpdateGuildValue("leader", Leader)
        Exit Sub
SetLeader_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetLeader", Erl)
End Sub

Public Function GetLeader() As Long
        On Error GoTo GetLeader_Err
100     GetLeader = p_Leader
        Exit Function
GetLeader_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.GetLeader", Erl)
End Function

Private Sub ReloadMemberList()
    On Error GoTo ReloadMemberList_Err
100     If MemberListLoaded Then Exit Sub
        Dim RS As ADODB.Recordset
102     Set RS = Query("SELECT user_id from guild_members where guild_id = ?", p_GuildId)
    
104     If Not RS.RecordCount = 0 Then
106         ReDim MemberList(RS.RecordCount - 1) As Long
            Dim i As Integer
108         i = 0
110         While Not RS.EOF
112             MemberList(i) = RS!user_id
                i = i + 1
                RS.MoveNext
            Wend
        End If
114     MemberCount = UBound(MemberList) + 1
116     MemberListLoaded = True
        Exit Sub
ReloadMemberList_Err:
118    Call TraceError(Err.Number, Err.Description, "clsClan.ReloadMemberList", Erl)
End Sub

Public Function GetMemberList() As Long()
    On Error GoTo GetMemberList_Err
        Call ReloadMemberList
        If MemberCount < (UBound(MemberList) + 1) Then
            Dim Members() As Long
            ReDim Members(MemberCount - 1) As Long
            Dim i As Integer
            For i = 0 To MemberCount - 1
                Members(i) = MemberList(i)
                GetMemberList = Members
                Exit Function
            Next i
        Else
110         GetMemberList = MemberList
        End If
        Exit Function
GetMemberList_Err:
112     Call TraceError(Err.Number, Err.Description, "clsClan.GetMemberList", Erl)
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
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
End Sub

Public Sub DesConectarMiembro(ByVal UserIndex As Integer)
        On Error GoTo DesConectarMiembro_Err

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
End Sub

Public Sub AceptarNuevoMiembro(ByRef UserId As Long)
    On Error GoTo AceptarNuevoMiembro_Err
100     If IsValidUserId(UserId) Then
102         Call SaveUserGuildIndex(UserId, p_GuildNumber)
104         Call SaveUserGuildAspirant(UserId, 0)
            Call Query("INSERT INTO guild_members (guild_id, user_id) VALUES (?,?)", p_GuildId, UserId)
            If Not MemberListLoaded Then
                Call ReloadMemberList
            Else
                MemberCount = MemberCount + 1
                ReDim Preserve MemberList(MemberCount - 1) As Long
                MemberList(MemberCount - 1) = UserId
            End If
        End If
        Exit Sub
AceptarNuevoMiembro_Err:
118     Call TraceError(Err.Number, Err.Description, "clsClan.AceptarNuevoMiembro", Erl)
End Sub

Public Sub ExpulsarMiembro(ByVal CharId As Long)
On Error GoTo ExpulsarMiembro_Err
        Call Execute("DELETE FROM guild_members WHERE user_id = ? AND guild_id = ?;", CharId, p_GuildId)
        Dim i As Integer
        For i = 0 To UBound(MemberList)
            If MemberList(i) = CharId Then
                MemberList(i) = 0
                Exit For
            End If
        Next i
        For i = i To UBound(MemberList) - 1
            MemberList(i) = MemberList(i + 1)
        Next i
        MemberCount = MemberCount - 1
        Call SaveUserGuildMember(CharId, p_GuildName)
        Call SaveUserGuildIndex(CharId, 0)
        Exit Sub
ExpulsarMiembro_Err:
142     Call TraceError(Err.Number, Err.Description, "clsClan.ExpulsarMiembro", Erl)
End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()
        
    On Error GoTo GetAspirantes_Err
        Dim RS As ADODB.Recordset
        Dim list() As String
102     Set RS = Query("SELECT user_id from guild_request where guild_id = ?", p_GuildId)
    
104     If Not RS.RecordCount = 0 Then
106         ReDim list(RS.RecordCount) As String
            Dim i As Integer
108         i = 0
110         While Not RS.EOF
112             list(i) = GetUserName(RS!user_id)
                i = i + 1
                RS.MoveNext
            Wend
        Else
            ReDim list(0) As String
        End If
    
114     GetAspirantes = list
        Exit Function
GetAspirantes_Err:
116     Call TraceError(Err.Number, Err.Description, "clsClan.GetAspirantes", Erl)
End Function

Public Function HasGuildRequest(ByVal UserId As Long) As Boolean
    On Error GoTo HasGuildRequest_Err
        Dim RS As ADODB.Recordset
        Dim list() As String
102     Set RS = Query("SELECT user_id from guild_request where guild_id = ? and user_id = ?", p_GuildId, UserId)
        HasGuildRequest = RS.RecordCount > 0
        Exit Function
HasGuildRequest_Err:
116     Call TraceError(Err.Number, Err.Description, "clsClan.HasGuildRequest", Erl)
End Function
Public Function CantidadAspirantes() As Integer
        On Error GoTo CantidadAspirantes_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT COUNT(*) FROM guild_request where guild_id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
        CantidadAspirantes = RS.Fields(0).Value
        Exit Function
CantidadAspirantes_Err:
108     Call TraceError(Err.Number, Err.Description, "clsClan.CantidadAspirantes", Erl)
End Function

Public Function DetallesSolicitudAspirante(ByVal username As String) As String
    On Error GoTo DetallesSolicitudAspirante_Err
        Dim CharId As Long
        CharId = GetCharacterIdWithName(username)
        If CharId <= 0 Then
            Exit Function
        End If
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT description FROM guild_request where guild_id = ? and user_id = ?", p_GuildId, CharId)
102     If RS Is Nothing Then Exit Function
        DetallesSolicitudAspirante = RS.Fields(0).Value
        Exit Function
DetallesSolicitudAspirante_Err:
104     Call TraceError(Err.Number, Err.Description, "clsClan.DetallesSolicitudAspirante", Erl)
End Function

Public Sub NuevoAspirante(ByRef name As String, ByRef Peticion As String)
        On Error GoTo NuevoAspirante_Err
        Dim CharId As Long
        CharId = GetCharacterIdWithName(name)
        If CharId <= 0 Then
            Exit Sub
        End If
        Call Execute("insert into guild_request (guild_id, user_id, description) values ( ?, ?, ?)", p_GuildId, CharId, Peticion)
        Exit Sub
NuevoAspirante_Err:
122     Call TraceError(Err.Number, Err.Description, "clsClan.NuevoAspirante", Erl)
End Sub

Public Sub RetirarAspirante(ByRef name As String)
        
        On Error GoTo RetirarAspirante_Err
        Dim CharId As Long
        CharId = GetCharacterIdWithName(name)
        If CharId <= 0 Then
            Exit Sub
        End If
        Call Execute("DELETE FROM guild_request WHERE user_id = ? AND guild_id = ?;", CharId, p_GuildId)
108     Call SaveUserGuildAspirant(CharId, 0)
        Call SaveUserGuildPedidos(name, p_GuildName)
        
        Exit Sub
RetirarAspirante_Err:
136     Call TraceError(Err.Number, Err.Description, "clsClan.RetirarAspirante", Erl)
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As String
        
    On Error GoTo GetFechaFundacion_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT strftime('%d/%m/%Y', creation_date, 'unixepoch') AS formatted_date FROM guilds where id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
104     GetFechaFundacion = RS.Fields(0).Value
        Exit Function
GetFechaFundacion_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetFechaFundacion", Erl)
End Function

Public Sub SetURL(ByRef URL As String)
        On Error GoTo SetURL_Err
100     Call ReplaceInvalidChars(URL)
110     Call UpdateGuildValue("url", URL)
        Exit Sub

SetURL_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetURL", Erl)
End Sub

Public Function GetURL() As String
    On Error GoTo GetURL_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT url FROM guilds where id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
104     GetURL = RS.Fields(0).Value
        Exit Function
GetURL_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetURL", Erl)
End Function

Public Sub SetNivelDeClan(ByRef nivel As Byte)
    On Error GoTo SetNivelDeClan_Err
        Call UpdateGuildValue("level", nivel)
        Exit Sub
SetNivelDeClan_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetNivelDeClan", Erl)
End Sub

Public Function GetNivelDeClan() As Byte
    On Error GoTo GetNivelDeClan_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT level FROM guilds where id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
104     GetNivelDeClan = RS.Fields(0).Value
        Exit Function
GetNivelDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetNivelDeClan", Erl)
End Function

Public Sub SetExpActual(ByRef Exp As Integer)
    On Error GoTo SetExpActual_Err
        Call UpdateGuildValue("current_exp", Exp)
    Exit Sub
SetExpActual_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.SetExpActual", Erl)
End Sub

Public Function GetExpActual() As Integer
    On Error GoTo GetExpActual_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT current_exp FROM guilds where id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
104     GetExpActual = RS.Fields(0).Value
        Exit Function
GetExpActual_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetExpActual", Erl)
End Function

Public Sub SetGuildNews(ByRef News As String)
    On Error GoTo SetGuildNews_Err
        
100     Call ReplaceInvalidChars(News)
102     News = Left$(News, NEWSLENGTH)
104     Call UpdateGuildValue("news", News)
        Exit Sub
SetGuildNews_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.SetGuildNews", Erl)
End Sub

Public Function GetGuildNews() As String
    On Error GoTo GetGuildNews_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT news FROM guilds where id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
104     GetGuildNews = RS.Fields(0).Value
        Exit Function
GetGuildNews_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetGuildNews", Erl)
End Function

Public Sub SetDesc(ByRef Desc As String)
        On Error GoTo SetDesc_Err
100     Call ReplaceInvalidChars(Desc)
102     Desc = Left$(Desc, DESCLENGTH)
104     Call UpdateGuildValue("description", Desc)
        Exit Sub
SetDesc_Err:
106     Call TraceError(Err.Number, Err.Description, "clsClan.SetDesc", Erl)
End Sub

Public Function GetDesc() As String
        
    On Error GoTo GetDesc_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT description FROM guilds where id = ?", p_GuildId)
102     If RS Is Nothing Then Exit Function
104     GetDesc = RS.Fields(0).Value
        Exit Function
GetDesc_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetDesc", Erl)
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
End Function


'ADMINISTRATIVAS
Public Sub ConectarGM(ByVal UserIndex As Integer)
        On Error GoTo ConectarGM_Err
100     p_GMsOnline.Add UserIndex
        Exit Sub
ConectarGM_Err:
102     Call TraceError(Err.Number, Err.Description, "clsClan.ConectarGM", Erl)
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

108     If InStrB(s, "¬") <> 0 Then
110         s = Replace(s, "¬", vbNullString)   'morgo usaba esto como "separador"
        End If
        Exit Sub
ReplaceInvalidChars_Err:
112     Call TraceError(Err.Number, Err.Description, "clsClan.ReplaceInvalidChars", Erl)
End Sub

