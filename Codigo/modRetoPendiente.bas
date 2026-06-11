Attribute VB_Name = "modPendingChallenge"
Option Explicit

Public Type t_PendingChallengePlayer
    Name As String
    Accepted As Boolean
    CurIndex As t_UserReference
End Type

Public Type t_PendingChallenge
    Challenger As t_UserReference
    Players() As Integer
    Bet As Long
    MaxPotions As Integer
    ItemsDrop As Boolean
End Type

Public PendingChallenges() As t_PendingChallenge
Public TotalPendingChallenges As Integer

Public Function NewPendingChallenge(ByVal Challenger As Integer, ByVal Bet As Long, _
                                    ByVal MaxPotions As Integer, ByVal ItemsDrop As Boolean) As Integer
    If TotalPendingChallenges = 0 Then
        ReDim PendingChallenges(1 To 1)
        TotalPendingChallenges = 1
    Else
        TotalPendingChallenges = TotalPendingChallenges + 1
        ReDim Preserve PendingChallenges(1 To TotalPendingChallenges)
    End If
    With PendingChallenges(TotalPendingChallenges)
        Call SetUserRef(.Challenger, Challenger)
        .Bet = Bet
        .MaxPotions = MaxPotions
        .ItemsDrop = ItemsDrop
    End With
    NewPendingChallenge = TotalPendingChallenges
End Function

Public Sub AddPlayerToChallenge(ByVal ChallengeIndex As Integer, ByVal UserIndex As Integer)
    Dim n As Integer
    With PendingChallenges(ChallengeIndex)
        If (Not Not .Players) = 0 Then
            ReDim .Players(1 To 1)
            n = 1
        Else
            n = UBound(.Players) + 1
            ReDim Preserve .Players(1 To n)
        End If
        .Players(n) = UserIndex
    End With
End Sub

Public Function GetPlayerIndexInPendingChallenge(ByVal UserIndex As Integer, ByVal ChallengeIndex As Integer) As Integer
    Dim i As Integer
    GetPlayerIndexInPendingChallenge = -1
    
    If ChallengeIndex <= 0 Then Exit Function
    
    With PendingChallenges(ChallengeIndex)
        For i = LBound(.Players) To UBound(.Players)
            If .Players(i) = UserIndex Then
                GetPlayerIndexInPendingChallenge = i
                Exit Function
            End If
        Next i
    End With
End Function
