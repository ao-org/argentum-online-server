Attribute VB_Name = "Unit_GmSupportHistory"
Option Explicit
#If UNIT_TEST = 1 Then

Public Function test_suite_gm_support_history() As Boolean
    On Error GoTo test_suite_gm_support_history_Err
    Call UnitTesting.RunTest("GM support history retains the newest replies", test_retains_newest_replies())
    Call UnitTesting.RunTest("GM support history preserves pending ordinary messages", test_preserves_ordinary_messages())
    Call UnitTesting.RunTest("GM support history extraction excludes ordinary messages", test_extracts_only_replies())
    Call UnitTesting.RunTest("Character saves do not overwrite GM reply history", test_character_save_preserves_replies())
    Call UnitTesting.RunTest("GM support normalizes plus-separated names", test_normalizes_plus_separated_names())
    Call UnitTesting.RunTest("GM support preserves multiline replies", test_multiline_reply_round_trip())
    test_suite_gm_support_history = True
    Exit Function
test_suite_gm_support_history_Err:
    Call UnitTesting.RunTestError("GM support history suite", Err.Description)
End Function

Private Function test_retains_newest_replies() As Boolean
    Dim source As String
    Dim limited As String
    Dim i As Long
    For i = 1 To UserMod.GM_REPLY_HISTORY_LIMIT + 2
        source = source & UserMod.GM_REPLY_MESSAGE_PREFIX & "GM" & Chr$(172) & "reply-" & CStr(i) & vbNewLine
    Next i
    limited = UserMod.LimitGmReplyHistory(source)
    test_retains_newest_replies = InStr(1, limited, "reply-1" & vbNewLine, vbBinaryCompare) = 0 And _
        InStr(1, limited, "reply-2" & vbNewLine, vbBinaryCompare) = 0 And _
        InStr(1, limited, "reply-3" & vbNewLine, vbBinaryCompare) > 0 And _
        InStr(1, limited, "reply-22" & vbNewLine, vbBinaryCompare) > 0
End Function

Private Function test_preserves_ordinary_messages() As Boolean
    Dim source As String
    Dim limited As String
    source = "ordinary notification" & vbNewLine & _
        UserMod.GM_REPLY_MESSAGE_PREFIX & "GM" & Chr$(172) & "answer" & vbNewLine
    limited = UserMod.LimitGmReplyHistory(source)
    test_preserves_ordinary_messages = InStr(1, limited, "ordinary notification" & vbNewLine, vbBinaryCompare) > 0 And _
        InStr(1, limited, UserMod.GM_REPLY_MESSAGE_PREFIX & "GM", vbBinaryCompare) > 0
End Function

Private Function test_extracts_only_replies() As Boolean
    Dim source As String
    Dim history As String
    source = "ordinary notification" & vbNewLine & _
        UserMod.GM_REPLY_MESSAGE_PREFIX & "GM" & Chr$(172) & "answer" & vbNewLine
    history = UserMod.ExtractGmReplyHistory(source)
    test_extracts_only_replies = InStr(1, history, "ordinary notification", vbBinaryCompare) = 0 And _
        InStr(1, history, UserMod.GM_REPLY_MESSAGE_PREFIX & "GM", vbBinaryCompare) > 0
End Function

Private Function test_character_save_preserves_replies() As Boolean
    test_character_save_preserves_replies = _
        InStr(1, QUERY_UPDATE_MAINPJ, "message_info", vbTextCompare) = 0 And _
        InStr(1, QUERY_UPDATE_MAINPJ_LOGOUT, "message_info", vbTextCompare) = 0
End Function

Private Function test_normalizes_plus_separated_names() As Boolean
    test_normalizes_plus_separated_names = _
        UserMod.NormalizeGmReplyUserName("John+Doe") = "John Doe" And _
        UserMod.NormalizeGmReplyUserName("John Doe") = "John Doe"
End Function

Private Function test_multiline_reply_round_trip() As Boolean
    Dim originalText  As String
    Dim encodedText   As String
    Dim storedHistory As String
    originalText = "first line %0A" & vbNewLine & "second line 100%"
    encodedText = UserMod.EncodeGmReplyMessage(originalText)
    storedHistory = UserMod.GM_REPLY_MESSAGE_PREFIX & "GM" & Chr$(172) & encodedText & vbNewLine
    test_multiline_reply_round_trip = _
        InStr(1, encodedText, vbCr, vbBinaryCompare) = 0 And _
        InStr(1, encodedText, vbLf, vbBinaryCompare) = 0 And _
        UserMod.DecodeGmReplyMessage(encodedText) = originalText And _
        UserMod.ExtractGmReplyHistory(storedHistory) = storedHistory
End Function

#End If
