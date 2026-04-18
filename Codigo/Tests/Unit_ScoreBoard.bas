Attribute VB_Name = "Unit_ScoreBoard"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' ScoreBoard Test Suite
' Tests initialization, add/retrieve, ranking order, and reset behavior.
' ==========================================================================

Public Function test_suite_scoreboard() As Boolean
    ' Example-based tests (Req 12.1, 12.2, 12.4)
    Call UnitTesting.RunTest("test_scoreboard_init_zero_entries", test_scoreboard_init_zero_entries())
    Call UnitTesting.RunTest("test_scoreboard_add_and_retrieve", test_scoreboard_add_and_retrieve())
    Call UnitTesting.RunTest("test_scoreboard_reset_clears_all", test_scoreboard_reset_clears_all())
    
    ' Property test (Property 15)
    Call UnitTesting.RunTest("test_prop_scoreboard_ranking_order", test_prop_scoreboard_ranking_order())
    
    test_suite_scoreboard = True
End Function

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' Req 12.1: A newly created ScoreBoard has zero entries.
' WHEN a ScoreBoard is initialized, THE ScoreBoard class SHALL start with
' zero entries.
Private Function test_scoreboard_init_zero_entries() As Boolean
    On Error GoTo Fail
    
    Dim board As New ScoreBoard
    
    ' GetRanking with size 5 should return entries all with Score = 0
    Dim ranking() As e_Rank
    ranking = board.GetRanking(5)
    
    Dim i As Integer
    For i = 0 To UBound(ranking)
        If ranking(i).Score <> 0 Then
            test_scoreboard_init_zero_entries = False
            Exit Function
        End If
        If ranking(i).PlayerIndex <> 0 Then
            test_scoreboard_init_zero_entries = False
            Exit Function
        End If
    Next i
    
    test_scoreboard_init_zero_entries = True
    Exit Function
Fail:
    test_scoreboard_init_zero_entries = False
End Function

' Req 12.2: Adding a score for a player makes it retrievable.
' WHEN a score is added for a player, THE ScoreBoard class SHALL record it
' and make it retrievable.
Private Function test_scoreboard_add_and_retrieve() As Boolean
    On Error GoTo Fail
    
    Dim board As New ScoreBoard
    
    ' Add a player and give them a score
    Call board.AddPlayer(1)
    Dim updatedScore As Integer
    updatedScore = board.UpdatePlayerScore(1, 10)
    
    If updatedScore <> 10 Then
        test_scoreboard_add_and_retrieve = False
        Exit Function
    End If
    
    ' Retrieve via GetRanking and verify the player appears with correct score
    Dim ranking() As e_Rank
    ranking = board.GetRanking(5)
    
    test_scoreboard_add_and_retrieve = (ranking(0).PlayerIndex = 1 And ranking(0).Score = 10)
    Exit Function
Fail:
    test_scoreboard_add_and_retrieve = False
End Function

' Req 12.4: Resetting the ScoreBoard clears all entries.
' WHEN the ScoreBoard is reset, THE ScoreBoard class SHALL clear all entries.
' Note: ScoreBoard has no explicit Reset method. Creating a new instance
' effectively resets state since the internal Dictionaries are re-initialized.
Private Function test_scoreboard_reset_clears_all() As Boolean
    On Error GoTo Fail
    
    Dim board As New ScoreBoard
    
    ' Add players and scores
    Call board.AddPlayer(1)
    Call board.UpdatePlayerScore(1, 50)
    Call board.AddPlayer(2)
    Call board.UpdatePlayerScore(2, 30)
    
    ' Verify scores exist
    Dim ranking() As e_Rank
    ranking = board.GetRanking(5)
    If ranking(0).Score = 0 Then
        test_scoreboard_reset_clears_all = False
        Exit Function
    End If
    
    ' "Reset" by creating a new instance
    Set board = New ScoreBoard
    
    ' Verify all entries are cleared
    ranking = board.GetRanking(5)
    
    Dim i As Integer
    For i = 0 To UBound(ranking)
        If ranking(i).Score <> 0 Or ranking(i).PlayerIndex <> 0 Then
            test_scoreboard_reset_clears_all = False
            Exit Function
        End If
    Next i
    
    test_scoreboard_reset_clears_all = True
    Exit Function
Fail:
    test_scoreboard_reset_clears_all = False
End Function

' --------------------------------------------------------------------------
' Property tests
' --------------------------------------------------------------------------

' Feature: unit-test-coverage, Property 15: ScoreBoard ranking order
' **Validates: Requirements 12.3**
'
' Loop over 100+ sets of random scores, add to ScoreBoard, verify
' GetRanking returns entries sorted in descending order by score.
Private Function test_prop_scoreboard_ranking_order() As Boolean
    On Error GoTo Fail
    
    Dim iteration As Long
    Dim numPlayers As Integer
    Dim j As Integer
    Dim score As Integer
    Dim board As ScoreBoard
    Dim ranking() As e_Rank
    
    For iteration = 1 To 110
        Set board = New ScoreBoard
        
        ' Add 5 to 10 players with deterministic "random" scores
        numPlayers = 5 + CInt(iteration Mod 6)
        
        For j = 1 To numPlayers
            Call board.AddPlayer(CLng(j))
            ' Deterministic pseudo-random score: (iteration * 7 + j * 13) Mod 1000
            score = CInt((iteration * 7 + j * 13) Mod 1000)
            If score > 0 Then
                Call board.UpdatePlayerScore(CInt(j), score)
            End If
        Next j
        
        ' Retrieve ranking
        ranking = board.GetRanking(numPlayers)
        
        ' Verify descending order: each score >= next score
        ' Only check entries that have been filled (Score > 0 or PlayerIndex > 0)
        Dim k As Integer
        For k = 0 To UBound(ranking) - 1
            ' Once we hit a zero-score entry, remaining should also be zero
            If ranking(k).Score = 0 Then
                ' All subsequent must also be 0
                Dim m As Integer
                For m = k + 1 To UBound(ranking)
                    If ranking(m).Score <> 0 Then
                        test_prop_scoreboard_ranking_order = False
                        Exit Function
                    End If
                Next m
                Exit For
            End If
            
            ' Current score must be >= next score (descending order)
            If ranking(k).Score < ranking(k + 1).Score Then
                test_prop_scoreboard_ranking_order = False
                Exit Function
            End If
        Next k
        
        Set board = Nothing
    Next iteration
    
    test_prop_scoreboard_ranking_order = True
    Exit Function
Fail:
    test_prop_scoreboard_ranking_order = False
End Function

#End If
