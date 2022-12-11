VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScoreBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim PlayersScore As New Dictionary
Dim DisconnectedUser As New Dictionary

Public Sub AddPlayer(ByVal PlayerIndex As Long)
    If Not PlayersScore.Exists(PlayerIndex) Then
        Call PlayersScore.Add(PlayerIndex, 0)
    End If
End Sub

Public Function UpdatePlayerScore(ByVal PlayerIndex As Integer, ByVal modifyValue As Integer) As Integer
    UpdatePlayerScore = max(PlayersScore.Item(PlayerIndex) + modifyValue, 0)
    PlayersScore.Item(PlayerIndex) = UpdatePlayerScore
End Function

Friend Function GetRanking(ByVal size As Integer) As e_Rank()
    ReDim rank(size) As e_Rank
    Dim key As Variant
    Dim pos As Integer
    Dim value As Integer
    For Each key In PlayersScore.Keys
        value = PlayersScore.Item(key)
        pos = InsertionPos(value, rank)
        If pos < size Then
            Call ShiftElement(pos, rank)
            rank(pos).PlayerIndex = key
            rank(pos).Score = value
        End If
    Next key
    GetRanking = rank
End Function

Friend Function GetTeamRanking(ByRef Lobby As t_Lobby) As e_Rank()
    ReDim rank(Lobby.RegisteredPlayers / Lobby.TeamSize) As e_Rank
    Dim key As Variant
    Dim i As Integer
    For i = i To UBound(rank)
        rank(i).PlayerIndex = i
    Next i
    For Each key In PlayersScore.Keys
        For i = 0 To Lobby.RegisteredPlayers - 1
            If Lobby.Players(i).user.ArrayIndex = key Then
                rank(Lobby.Players(i).Team).Score = rank(Lobby.Players(i).Team).Score + PlayersScore.Item(key)
            End If
        Next i
    Next key
    Call SortRank(rank)
    GetTeamRanking = rank
End Function

Private Function InsertionPos(ByVal Score As Integer, ByRef rank() As e_Rank) As Integer
    If rank(UBound(rank)).Score >= Score Then
        InsertionPos = UBound(rank) + 1
        Exit Function
    End If
    
    Dim currentPosition As Integer
    For currentPosition = 0 To UBound(rank)
        If rank(currentPosition).Score < Score Then
            Exit For
        End If
    Next currentPosition
    InsertionPos = currentPosition
End Function

Private Sub SortRank(ByRef rank() As e_Rank)
    Dim CurrentMax As Integer
    Dim currentPos As Integer
    Dim Temp As e_Rank
    Dim i As Integer
    Dim j As Integer
    For i = 1 To UBound(rank)
        currentPos = i
        CurrentMax = rank(currentPos).Score
        For j = i + 1 To UBound(rank)
            If rank(j).Score > CurrentMax Then
                CurrentMax = rank(j).Score
                currentPos = j
            End If
        Next j
        If CurrentMax > rank(i).Score Then
            Temp = rank(i)
            rank(i) = rank(currentPos)
            rank(currentPos) = Temp
        End If
    Next i
    
End Sub

Private Sub ShiftElement(ByVal position As Integer, ByRef rank() As e_Rank)
    Dim i As Integer
    For i = UBound(rank) To position + 1 Step -1
        rank(i) = rank(i - 1)
    Next i
End Sub

Friend Sub PublishScoreboard(ByVal Audience As Integer, ByVal AudienceIndex, ByRef Lobby As t_Lobby, ByVal IsFinal As Boolean)
    Dim i As Integer
    Dim rank() As e_Rank
    If IsFinal And Lobby.TeamSize > 0 Then
        Call PublishTeamScore(Audience, AudienceIndex, Lobby)
        Exit Sub
    End If
    If Lobby.TeamSize >= 1 Then
        rank = GetTeamRanking(Lobby)
        Call SendData(Audience, AudienceIndex, PrepareMessageConsoleMsg("Posiciones:", e_FontTypeNames.FONTTYPE_GUILD))
        For i = 1 To UBound(rank)
            Call SendData(Audience, AudienceIndex, PrepareMessageConsoleMsg("Equipo " & (rank(i).PlayerIndex) & ": " & rank(i).Score, e_FontTypeNames.FONTTYPE_GUILD))
        Next i
    Else
        rank = GetRanking(5)
        For i = 0 To UBound(rank) - 1
            If rank(i).Score <= 0 Then
                Exit Sub
            End If
            If i = 0 Then
                Call SendData(Audience, AudienceIndex, PrepareMessageConsoleMsg("Posiciones:", e_FontTypeNames.FONTTYPE_GUILD))
            End If
            Call SendData(Audience, AudienceIndex, PrepareMessageConsoleMsg((i + 1) & ") " & UserList(rank(i).PlayerIndex).name & ": " & rank(i).Score, e_FontTypeNames.FONTTYPE_GUILD))
        Next i
    End If
End Sub

Friend Sub PublishTeamScore(ByVal Audience As Integer, ByVal AudienceIndex, ByRef Lobby As t_Lobby)
    Dim rank() As e_Rank
    rank = GetTeamRanking(Lobby)
    Dim Winner As Integer
    Winner = rank(1).PlayerIndex
    Dim TeamPlayerList As String
    Dim key As Variant
    Dim i As Integer
    Call SendData(Audience, AudienceIndex, PrepareMessageConsoleMsg("Ganador equipo " & Winner & ": " & rank(1).Score, e_FontTypeNames.FONTTYPE_GUILD))
    For Each key In PlayersScore.Keys
        For i = 0 To Lobby.RegisteredPlayers - 1
            If Lobby.Players(i).user.ArrayIndex = key And Lobby.Players(i).Team = Winner Then
                TeamPlayerList = TeamPlayerList & UserList(key).name & " (" & PlayersScore.Item(key) & ") "
            End If
        Next i
    Next key
    Call SendData(Audience, AudienceIndex, PrepareMessageConsoleMsg(TeamPlayerList, e_FontTypeNames.FONTTYPE_GUILD))
End Sub

Public Sub DisconnectUser(ByVal userIndex As Integer, ByVal userID As Long)
    If Not PlayersScore.Exists(userIndex) Then
        Exit Sub
    End If
    Call DisconnectedUser.Add(userID, PlayersScore.Item(userIndex))
    PlayersScore.Remove (userIndex)
End Sub

Public Sub RestoreDisconnectUser(ByVal userIndex As Integer, ByVal userID As Long)

    If Not DisconnectedUser.Exists(userID) Then
        Exit Sub
    End If
    If PlayersScore.Exists(userIndex) Then
        PlayersScore.Item(userIndex) = DisconnectedUser.Item(userID)
    Else
        Call PlayersScore.Add(userIndex, DisconnectedUser.Item(userID))
    End If
    DisconnectedUser.Remove (userID)
End Sub
