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

Dim PlayersScore As New Dictionary
Dim DisconnectedUser As New Dictionary

Public Sub AddPlayer(ByVal PlayerIndex As Long)
    On Error Goto AddPlayer_Err
    If Not PlayersScore.Exists(PlayerIndex) Then
        Call PlayersScore.Add(PlayerIndex, 0)
    End If
    Exit Sub
AddPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.AddPlayer", Erl)
End Sub

Public Function UpdatePlayerScore(ByVal PlayerIndex As Integer, ByVal modifyValue As Integer) As Integer
    On Error Goto UpdatePlayerScore_Err
    If Not PlayersScore.Exists(PlayerIndex) Then Exit Function
    UpdatePlayerScore = max(PlayersScore.Item(PlayerIndex) + modifyValue, 0)
    PlayersScore.Item(PlayerIndex) = UpdatePlayerScore
    Exit Function
UpdatePlayerScore_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.UpdatePlayerScore", Erl)
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
    On Error Goto InsertionPos_Err
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
    Exit Function
InsertionPos_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.InsertionPos", Erl)
End Function

Private Sub SortRank(ByRef rank() As e_Rank)
    On Error Goto SortRank_Err
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
    Exit Sub
SortRank_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.SortRank", Erl)
End Sub

Private Sub ShiftElement(ByVal position As Integer, ByRef rank() As e_Rank)
    On Error Goto ShiftElement_Err
    Dim i As Integer
    For i = UBound(rank) To position + 1 Step -1
        rank(i) = rank(i - 1)
    Next i
    Exit Sub
ShiftElement_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.ShiftElement", Erl)
End Sub

Friend Sub PublishScoreboard(ByVal Audience As Integer, ByVal AudienceIndex, ByRef Lobby As t_Lobby, ByVal IsFinal As Boolean)
    Dim i As Integer
    Dim rank() As e_Rank
    If IsFinal And Lobby.TeamSize > 1 Then
        Call PublishTeamScore(Audience, AudienceIndex, Lobby)
        Exit Sub
    End If
    If Lobby.TeamSize > 1 Then
        rank = GetTeamRanking(Lobby)
        Call SendData(Audience, AudienceIndex, PrepareMessageLocaleMsg(1838, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1838=Posiciones:
        For i = 1 To UBound(rank)
            Call SendData(Audience, AudienceIndex, PrepareMessageLocaleMsg(1839, rank(i).PlayerIndex & "¬" & rank(i).Score, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1839=Equipo ¬1: ¬2
        Next i
    Else
        rank = GetRanking(5)
        For i = 0 To UBound(rank) - 1
            If rank(i).Score <= 0 Then
                Exit Sub
            End If
            If i = 0 Then
                Call SendData(Audience, AudienceIndex, PrepareMessageLocaleMsg(1838, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1838=Posiciones:
            End If
            Call SendData(Audience, AudienceIndex, PrepareMessageLocaleMsg(1824, (i + 1) & "¬" & UserList(rank(i).PlayerIndex).name & "¬" & rank(i).Score & "¬", e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1824=¬1) ¬2: ¬3 puntos¬4
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
    Call SendData(Audience, AudienceIndex, PrepareMessageLocaleMsg(1840, Winner & "¬" & rank(1).Score, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1840=Ganador equipo ¬1: ¬2

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
    On Error Goto DisconnectUser_Err
    If Not PlayersScore.Exists(userIndex) Then
        Exit Sub
    End If
    Call DisconnectedUser.Add(userID, PlayersScore.Item(userIndex))
    PlayersScore.Remove (userIndex)
    Exit Sub
DisconnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.DisconnectUser", Erl)
End Sub

Public Sub RestoreDisconnectUser(ByVal userIndex As Integer, ByVal userID As Long)
    On Error Goto RestoreDisconnectUser_Err

    If Not DisconnectedUser.Exists(userID) Then
        Exit Sub
    End If
    If PlayersScore.Exists(userIndex) Then
        PlayersScore.Item(userIndex) = DisconnectedUser.Item(userID)
    Else
        Call PlayersScore.Add(userIndex, DisconnectedUser.Item(userID))
    End If
    DisconnectedUser.Remove (userID)
    Exit Sub
RestoreDisconnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "ScoreBoard.RestoreDisconnectUser", Erl)
End Sub
