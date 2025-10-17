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
Private p_IteradorOnlineGMs     As Integer
Private p_GuildNumber           As Integer      'Numero de guild en el mundo
Private p_GuildId               As Long
Private p_GuildLevel            As Integer
Private p_CurrentExp            As Integer
Private p_GuildNews             As String
Private Const NEWSLENGTH        As Integer = 1024
Private Const DESCLENGTH        As Integer = 256
Private MemberList()            As Long
Private memberCount             As Integer
Private MemberListLoaded        As Boolean

Public Property Get GuildName() As String
    On Error GoTo GuildName_Err
    GuildName = p_GuildName
    Exit Property
GuildName_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GuildName", Erl)
End Property

Private Sub UpdateGuildValue(ByVal ColumnName As String, ByVal value As Variant)
    Call Execute("UPDATE guilds SET " & ColumnName & " = ? WHERE id = ?;", value, p_GuildId)
End Sub

'
'ALINEACION Y ANTIFACCION
'
Public Property Get Alineacion() As e_ALINEACION_GUILD
    On Error GoTo Alineacion_Err
    Alineacion = p_Alineacion
    Exit Property
Alineacion_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.Alineacion", Erl)
End Property

Private Sub Class_Initialize()
    Set p_OnlineMembers = New Collection
    Set p_GMsOnline = New Collection
    p_IteradorOnlineMembers = 0
    p_IteradorOnlineGMs = 0
    MemberListLoaded = False
End Sub

Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Err
    Set p_OnlineMembers = Nothing
    Set p_GMsOnline = Nothing
    Exit Sub
Class_Terminate_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.Class_Terminate", Erl)
End Sub

Public Sub InitFromRecord(RS As Recordset, ByVal GuildNumber As Integer)
    On Error GoTo Inicializar_Err
    p_GuildName = RS!guild_name
    p_GuildId = RS!Id
    p_Alineacion = RS!alignment
    p_GuildNumber = GuildNumber
    p_Leader = RS!leader_id
    p_Founder = RS!founder_id
    Exit Sub
Inicializar_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.Inicializar", Erl)
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
        p_GuildId = val(RS.Fields(0).value)
    End If
    Exit Sub
InicializarNuevoClan_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.InicializarNuevoClan", Erl)
End Sub

Public Sub ProcesarFundacionDeOtroClan()
    On Error GoTo ProcesarFundacionDeOtroClan_Err
    Exit Sub
ProcesarFundacionDeOtroClan_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.ProcesarFundacionDeOtroClan", Erl)
End Sub

'
'MEMBRESIAS
'
Public Property Get Fundador() As Long
    On Error GoTo Fundador_Err
    Fundador = p_Founder
    Exit Property
Fundador_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.Fundador", Erl)
End Property

Public Property Get CantidadDeMiembros() As Integer
    On Error GoTo CantidadDeMiembros_Err
    Call ReloadMemberList
    CantidadDeMiembros = memberCount
    Exit Property
CantidadDeMiembros_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.CantidadDeMiembros", Erl)
End Property

Public Sub SetLeader(ByRef leader As Long)
    On Error GoTo SetLeader_Err
    Call UpdateGuildValue("leader", leader)
    Exit Sub
SetLeader_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetLeader", Erl)
End Sub

Public Function GetLeader() As Long
    On Error GoTo GetLeader_Err
    GetLeader = p_Leader
    Exit Function
GetLeader_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetLeader", Erl)
End Function

Private Sub ReloadMemberList()
    On Error GoTo ReloadMemberList_Err
    If MemberListLoaded Then Exit Sub
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT user_id from guild_members where guild_id = ?", p_GuildId)
    If Not RS.RecordCount = 0 Then
        ReDim MemberList(RS.RecordCount - 1) As Long
        Dim i As Integer
        i = 0
        While Not RS.EOF
            MemberList(i) = RS!user_id
            i = i + 1
            RS.MoveNext
        Wend
    End If
    memberCount = UBound(MemberList) + 1
    MemberListLoaded = True
    Exit Sub
ReloadMemberList_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.ReloadMemberList", Erl)
End Sub

Public Function GetMemberList() As Long()
    On Error GoTo GetMemberList_Err
    Call ReloadMemberList
    If memberCount < (UBound(MemberList) + 1) Then
        Dim Members() As Long
        ReDim Members(memberCount - 1) As Long
        Dim i As Integer
        For i = 0 To memberCount - 1
            Members(i) = MemberList(i)
        Next i
        GetMemberList = Members
    Else
        GetMemberList = MemberList
    End If
    Exit Function
GetMemberList_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetMemberList", Erl)
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
    On Error GoTo ConectarMiembro_Err
    p_OnlineMembers.Add UserIndex
    With UserList(UserIndex)
        ' No avisa cuando loguea un dios
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
            Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat("Msg2037¬" & .name, 9)) 'Msg2037=¬1 se ha conectado.
        End If
    End With
    Exit Sub
ConectarMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.ConectarMiembro", Erl)
End Sub

Public Sub DesConectarMiembro(ByVal UserIndex As Integer)
    On Error GoTo DesConectarMiembro_Err
    Dim i As Integer
    For i = 1 To p_OnlineMembers.count
        If p_OnlineMembers.Item(i) = UserIndex Then
            p_OnlineMembers.Remove i
            With UserList(UserIndex)
                ' No avisa cuando se desconecta un dios
                If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
                    Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat("Msg2038¬" & .name, 8)) 'Msg2038=¬1 se ha desconectado.
                End If
            End With
            Exit Sub
        End If
    Next i
    Exit Sub
DesConectarMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.DesConectarMiembro", Erl)
End Sub

Public Sub AceptarNuevoMiembro(ByRef UserId As Long)
    On Error GoTo AceptarNuevoMiembro_Err
    If IsValidUserId(UserId) Then
        Call SaveUserGuildIndex(UserId, p_GuildNumber)
        Call SaveUserGuildAspirant(UserId, 0)
        Call Query("INSERT INTO guild_members (guild_id, user_id) VALUES (?,?)", p_GuildId, UserId)
        If Not MemberListLoaded Then
            Call ReloadMemberList
        Else
            memberCount = memberCount + 1
            ReDim Preserve MemberList(memberCount - 1) As Long
            MemberList(memberCount - 1) = UserId
        End If
    End If
    Exit Sub
AceptarNuevoMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.AceptarNuevoMiembro", Erl)
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
    memberCount = memberCount - 1
    Call SaveUserGuildMember(CharId, p_GuildName)
    Call SaveUserGuildIndex(CharId, 0)
    Exit Sub
ExpulsarMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.ExpulsarMiembro", Erl)
End Sub

'
'ASPIRANTES
'
Public Function GetAspirantes() As String()
    On Error GoTo GetAspirantes_Err
    Dim RS     As ADODB.Recordset
    Dim list() As String
    Set RS = Query("SELECT user_id from guild_request where guild_id = ?", p_GuildId)
    If Not RS.RecordCount = 0 Then
        ReDim list(RS.RecordCount) As String
        Dim i As Integer
        i = 0
        While Not RS.EOF
            list(i) = GetUserName(RS!user_id)
            i = i + 1
            RS.MoveNext
        Wend
    Else
        ReDim list(0) As String
    End If
    GetAspirantes = list
    Exit Function
GetAspirantes_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetAspirantes", Erl)
End Function

Public Function HasGuildRequest(ByVal UserId As Long) As Boolean
    On Error GoTo HasGuildRequest_Err
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT user_id from guild_request where guild_id = ? and user_id = ?", p_GuildId, UserId)
    HasGuildRequest = RS.RecordCount > 0
    Exit Function
HasGuildRequest_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.HasGuildRequest", Erl)
End Function

Public Function CantidadAspirantes() As Integer
    On Error GoTo CantidadAspirantes_Err
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT COUNT(*) FROM guild_request where guild_id = ?", p_GuildId)
    If RS Is Nothing Then Exit Function
    CantidadAspirantes = RS.Fields(0).value
    Exit Function
CantidadAspirantes_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.CantidadAspirantes", Erl)
End Function

Public Function DetallesSolicitudAspirante(ByVal username As String) As String
    On Error GoTo DetallesSolicitudAspirante_Err
    Dim CharId As Long
    CharId = GetCharacterIdWithName(username)
    If CharId <= 0 Then
        Exit Function
    End If
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT description FROM guild_request where guild_id = ? and user_id = ?", p_GuildId, CharId)
    If RS Is Nothing Then Exit Function
    DetallesSolicitudAspirante = RS.Fields(0).value
    Exit Function
DetallesSolicitudAspirante_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.DetallesSolicitudAspirante", Erl)
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
    Call TraceError(Err.Number, Err.Description, "clsClan.NuevoAspirante", Erl)
End Sub

Public Sub RetirarAspirante(ByRef name As String)
    On Error GoTo RetirarAspirante_Err
    Dim CharId As Long
    CharId = GetCharacterIdWithName(name)
    If CharId <= 0 Then
        Exit Sub
    End If
    Call Execute("DELETE FROM guild_request WHERE user_id = ? AND guild_id = ?;", CharId, p_GuildId)
    Call SaveUserGuildAspirant(CharId, 0)
    Call SaveUserGuildPedidos(name, p_GuildName)
    Exit Sub
RetirarAspirante_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.RetirarAspirante", Erl)
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'
Public Function GetFechaFundacion() As String
    On Error GoTo GetFechaFundacion_Err
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT strftime('%d/%m/%Y', creation_date, 'unixepoch') AS formatted_date FROM guilds where id = ?", p_GuildId)
    If RS Is Nothing Then Exit Function
    GetFechaFundacion = RS.Fields(0).value
    Exit Function
GetFechaFundacion_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetFechaFundacion", Erl)
End Function

Public Sub SetURL(ByRef URL As String)
    On Error GoTo SetURL_Err
    Call ReplaceInvalidChars(URL)
    Call UpdateGuildValue("url", URL)
    Exit Sub
SetURL_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetURL", Erl)
End Sub

Public Sub SetNivelDeClan(ByRef nivel As Byte)
    On Error GoTo SetNivelDeClan_Err
    Call UpdateGuildValue("level", nivel)
    p_GuildLevel = nivel
    Exit Sub
SetNivelDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetNivelDeClan", Erl)
End Sub

Public Function GetNivelDeClan() As Integer
    If p_GuildLevel = 0 Then
        p_GuildLevel = GetStoredClanLevel
    End If
    GetNivelDeClan = p_GuildLevel
End Function

Public Function GetStoredClanLevel() As Byte
    On Error GoTo GetNivelDeClan_Err
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT level FROM guilds where id = ?", p_GuildId)
    If RS Is Nothing Then Exit Function
    GetStoredClanLevel = RS.Fields(0).value
    Exit Function
GetNivelDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetNivelDeClan", Erl)
End Function

Public Sub SetExpActual(ByRef Exp As Integer)
    On Error GoTo SetExpActual_Err
    Call UpdateGuildValue("current_exp", Exp)
    p_CurrentExp = Exp
    Exit Sub
SetExpActual_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetExpActual", Erl)
End Sub

Public Function GetExpActual() As Integer
    On Error GoTo GetExpActual_Err
    If p_CurrentExp = 0 Then
        Dim RS As ADODB.Recordset
        Set RS = Query("SELECT current_exp FROM guilds where id = ?", p_GuildId)
        If RS Is Nothing Then Exit Function
        p_CurrentExp = RS.Fields(0).value
    End If
    GetExpActual = p_CurrentExp
    Exit Function
GetExpActual_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetExpActual", Erl)
End Function

Public Sub SetGuildNews(ByRef News As String)
    On Error GoTo SetGuildNews_Err
    Call ReplaceInvalidChars(News)
    News = Left$(News, NEWSLENGTH)
    Call UpdateGuildValue("news", News)
    Exit Sub
SetGuildNews_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetGuildNews", Erl)
End Sub

Public Function GetGuildNews() As String
    On Error GoTo GetGuildNews_Err
    If p_GuildNews = "" Then
        Dim RS As ADODB.Recordset
        Set RS = Query("SELECT news FROM guilds where id = ?", p_GuildId)
        If RS Is Nothing Then Exit Function
        p_GuildNews = RS.Fields(0).value
    End If
    GetGuildNews = p_GuildNews
    Exit Function
GetGuildNews_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetGuildNews", Erl)
End Function

Public Sub SetDesc(ByRef Desc As String)
    On Error GoTo SetDesc_Err
    Call ReplaceInvalidChars(Desc)
    Desc = Left$(Desc, DESCLENGTH)
    Call UpdateGuildValue("description", Desc)
    Exit Sub
SetDesc_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.SetDesc", Erl)
End Sub

Public Function GetDesc() As String
    On Error GoTo GetDesc_Err
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT description FROM guilds where id = ?", p_GuildId)
    If RS Is Nothing Then Exit Function
    GetDesc = RS.Fields(0).value
    Exit Function
GetDesc_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.GetDesc", Erl)
End Function

Public Function m_Iterador_ProximoUserIndex() As Integer
    On Error GoTo m_Iterador_ProximoUserIndex_Err
    If p_IteradorOnlineMembers < p_OnlineMembers.count Then
        p_IteradorOnlineMembers = p_IteradorOnlineMembers + 1
        m_Iterador_ProximoUserIndex = p_OnlineMembers.Item(p_IteradorOnlineMembers)
    Else
        p_IteradorOnlineMembers = 0
        m_Iterador_ProximoUserIndex = 0
    End If
    Exit Function
m_Iterador_ProximoUserIndex_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.m_Iterador_ProximoUserIndex", Erl)
End Function

Public Function Iterador_ProximoGM() As Integer
    On Error GoTo Iterador_ProximoGM_Err
    If p_IteradorOnlineGMs < p_GMsOnline.count Then
        p_IteradorOnlineGMs = p_IteradorOnlineGMs + 1
        Iterador_ProximoGM = p_GMsOnline.Item(p_IteradorOnlineGMs)
    Else
        p_IteradorOnlineGMs = 0
        Iterador_ProximoGM = 0
    End If
    Exit Function
Iterador_ProximoGM_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.Iterador_ProximoGM", Erl)
End Function

'ADMINISTRATIVAS
Public Sub ConectarGM(ByVal UserIndex As Integer)
    On Error GoTo ConectarGM_Err
    p_GMsOnline.Add UserIndex
    Exit Sub
ConectarGM_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.ConectarGM", Erl)
End Sub

Public Sub DesconectarGM(ByVal UserIndex As Integer)
    On Error GoTo DesconectarGM_Err
    Dim i As Integer
    For i = 1 To p_GMsOnline.count
        If p_GMsOnline.Item(i) = UserIndex Then
            p_GMsOnline.Remove (i)
        End If
    Next i
    Exit Sub
DesconectarGM_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.DesconectarGM", Erl)
End Sub

'
'VARIAS, EXTRAS Y DEMASES
'
Private Sub ReplaceInvalidChars(ByRef s As String)
    On Error GoTo ReplaceInvalidChars_Err
    If InStrB(s, Chr$(13)) <> 0 Then
        s = Replace(s, Chr$(13), vbNullString)
    End If
    If InStrB(s, Chr$(10)) <> 0 Then
        s = Replace(s, Chr$(10), vbNullString)
    End If
    If InStrB(s, "¬") <> 0 Then
        s = Replace(s, "¬", vbNullString)   'morgo usaba esto como "separador"
    End If
    Exit Sub
ReplaceInvalidChars_Err:
    Call TraceError(Err.Number, Err.Description, "clsClan.ReplaceInvalidChars", Erl)
End Sub
