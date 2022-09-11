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

Private Type e_Rank
    PlayerIndex As Integer
    Score As Integer
End Type

Dim TopList() As e_Rank
Dim PlayersScore As New Dictionary

Public Sub Initiaize(ByVal rankSize As Integer)
    ReDim TopList(rankSize) As e_Rank
End Sub

Public Sub AddPlayer(ByVal PlayerIndex As Integer)
    If Not PlayersScore.Exists(PlayerIndex) Then
        Call PlayersScore.Add(PlayerIndex, 0)
    End If
End Sub

Public Function UpdatePlayerScore(ByVal PlayerIndex As Integer, ByVal modifyValue As Integer) As Integer
    UpdatePlayerScore = max(PlayersScore.Item(PlayerIndex) + modifyValue, 0)
    PlayersScore.Item(PlayerIndex) = UpdatePlayerScore
    Call UpdateRanking(PlayersScore.Item(PlayerIndex), PlayerIndex)
End Function

Private Sub UpdateRanking(ByVal value As Integer, ByVal PlayerIndex As Integer)
    Dim insertPos As Integer
    insertPos = InsertionPos(value)
    If (insertPos <= UBound(TopList)) Then
        If TopList(insertPos).PlayerIndex <> PlayerIndex Then
            Call ShiftElement(insertPos, CurrentScorePos(PlayerIndex))
        End If
        TopList(insertPos).PlayerIndex = PlayerIndex
        TopList(insertPos).Score = value
    End If
End Sub

Private Function InsertionPos(ByVal Score As Integer) As Integer
    If TopList(UBound(TopList)).Score >= Score Then
        InsertionPos = UBound(TopList) + 1
        Exit Function
    End If
    
    Dim currentPosition As Integer
    For currentPosition = 0 To UBound(TopList)
        If TopList(currentPosition).Score < Score Then
            Exit For
        End If
    Next currentPosition
    InsertionPos = currentPosition
End Function

Private Function CurrentScorePos(ByVal PlayerIndex As Integer) As Integer
    Dim i As Integer
    For i = 0 To UBound(TopList) - 1
        If TopList(i).PlayerIndex = PlayerIndex Then
            Exit For
        End If
    Next i
    CurrentScorePos = i
End Function

Private Sub ShiftElement(ByVal position As Integer, ByVal starPos As Integer)
    Dim i As Integer
    For i = starPos To position + 1 Step -1
        TopList(i) = TopList(i - 1)
    Next i
End Sub

Public Sub PublishScoreboard(ByVal audience As Integer, ByVal audienceIndex)
    Dim i As Integer
    For i = 0 To UBound(TopList) - 1
        If TopList(i).Score <= 0 Then
            Exit Sub
        End If
        If i = 0 Then
            Call SendData(audience, audienceIndex, PrepareMessageConsoleMsg("Posiciones:", e_FontTypeNames.FONTTYPE_GUILD))
        End If
        Call SendData(audience, audienceIndex, PrepareMessageConsoleMsg((i + 1) & ") " & UserList(TopList(i).PlayerIndex).name & ": " & TopList(i).Score, e_FontTypeNames.FONTTYPE_GUILD))
    Next i
End Sub
