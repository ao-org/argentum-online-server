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

Public Sub AddPlayer(ByVal PlayerIndex As Integer)
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

Private Sub ShiftElement(ByVal position As Integer, ByRef rank() As e_Rank)
    Dim i As Integer
    For i = UBound(rank) To position + 1 Step -1
        rank(i) = rank(i - 1)
    Next i
End Sub

Public Sub PublishScoreboard(ByVal audience As Integer, ByVal audienceIndex)
    Dim i As Integer
    Dim rank() As e_Rank
    rank = GetRanking(5)
    For i = 0 To UBound(rank) - 1
        If rank(i).Score <= 0 Then
            Exit Sub
        End If
        If i = 0 Then
            Call SendData(audience, audienceIndex, PrepareMessageConsoleMsg("Posiciones:", e_FontTypeNames.FONTTYPE_GUILD))
        End If
        Call SendData(audience, audienceIndex, PrepareMessageConsoleMsg((i + 1) & ") " & UserList(rank(i).PlayerIndex).name & ": " & rank(i).Score, e_FontTypeNames.FONTTYPE_GUILD))
    Next i
End Sub
