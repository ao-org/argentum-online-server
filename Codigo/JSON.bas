Attribute VB_Name = "mod_JSON"
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
'
' VBJSONDeserializer is a VB6 adaptation of the VB-JSON project @
' Fuente: https://www.codeproject.com/Articles/720368/VB-JSON-Parser-Improved-Performance
' BSD Licensed
Option Explicit
' DECLARACIONES API
Private Declare Function GetLocaleInfo _
                Lib "kernel32.dll" _
                Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "Kernel32" ()
' CONSTANTES LOCALE API
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SGROUPING = &H10
' CONSTANTES JSON
Private Const A_CURLY_BRACKET_OPEN   As Integer = 123       ' AscW("{")
Private Const A_CURLY_BRACKET_CLOSE  As Integer = 125       ' AscW("}")
Private Const A_SQUARE_BRACKET_OPEN  As Integer = 91        ' AscW("[")
Private Const A_SQUARE_BRACKET_CLOSE As Integer = 93        ' AscW("]")
Private Const A_BRACKET_OPEN         As Integer = 40        ' AscW("(")
Private Const A_BRACKET_CLOSE        As Integer = 41        ' AscW(")")
Private Const A_COMMA                As Integer = 44        ' AscW(",")
Private Const A_DOUBLE_QUOTE         As Integer = 34        ' AscW("""")
Private Const A_SINGLE_QUOTE         As Integer = 39        ' AscW("'")
Private Const A_BACKSLASH            As Integer = 92        ' AscW("\")
Private Const A_FORWARDSLASH         As Integer = 47        ' AscW("/")
Private Const A_COLON                As Integer = 58        ' AscW(":")
Private Const A_SPACE                As Integer = 32        ' AscW(" ")
Private Const A_ASTERIX              As Integer = 42        ' AscW("*")
Private Const A_VBCR                 As Integer = 13        ' AscW("vbcr")
Private Const A_VBLF                 As Integer = 10        ' AscW("vblf")
Private Const A_VBTAB                As Integer = 9         ' AscW("vbTab")
Private Const A_VBCRLF               As Integer = 13        ' AscW("vbcrlf")
Private Const A_b                    As Integer = 98        ' AscW("b")
Private Const A_f                    As Integer = 102       ' AscW("f")
Private Const A_n                    As Integer = 110       ' AscW("n")
Private Const A_r                    As Integer = 114       ' AscW("r"
Private Const A_t                    As Integer = 116       ' AscW("t"))
Private Const A_u                    As Integer = 117       ' AscW("u")
Private m_decSep                     As String
Private m_groupSep                   As String
Private m_parserrors                 As String
Private m_str()                      As Integer
Private m_length                     As Long

Public Function GetParserErrors() As String
    On Error GoTo GetParserErrors_Err
    GetParserErrors = m_parserrors
    Exit Function
GetParserErrors_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.GetParserErrors", Erl)
End Function

Public Function parse(ByRef str As String) As Object
    On Error GoTo parse_Err
    m_decSep = GetRegionalSettings(LOCALE_SDECIMAL)
    m_groupSep = GetRegionalSettings(LOCALE_SGROUPING)
    Dim Index As Long
    Index = 1
    Call GenerateStringArray(str)
    m_parserrors = vbNullString
    Call skipChar(Index)
    Select Case m_str(Index)
        Case A_SQUARE_BRACKET_OPEN
            Set parse = parseArray(str, Index)
        Case A_CURLY_BRACKET_OPEN
            Set parse = parseObject(str, Index)
        Case Else
            m_parserrors = "JSON Invalido"
    End Select
    'clean array
    ReDim m_str(1)
    Exit Function
parse_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parse", Erl)
End Function

Private Sub GenerateStringArray(ByRef str As String)
    On Error GoTo GenerateStringArray_Err
    Dim i As Long
    m_length = Len(str)
    ReDim m_str(1 To m_length)
    For i = 1 To m_length
        m_str(i) = AscW(mid$(str, i, 1))
    Next i
    Exit Sub
GenerateStringArray_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.GenerateStringArray", Erl)
End Sub

Private Function parseObject(ByRef str As String, ByRef Index As Long) As Dictionary
    On Error GoTo parseObject_Err
    Set parseObject = New Dictionary
    Dim sKey    As String
    Dim charint As Integer
    Call skipChar(Index)
    If m_str(Index) <> A_CURLY_BRACKET_OPEN Then
        m_parserrors = m_parserrors & "Objeto invalido en la posicion " & Index & " : " & mid$(str, Index) & vbCrLf
        Exit Function
    End If
    Index = Index + 1
    Do
        Call skipChar(Index)
        charint = m_str(Index)
        Select Case charint
            Case A_COMMA
                Index = Index + 1
                Call skipChar(Index)
            Case A_CURLY_BRACKET_CLOSE
                Index = Index + 1
                Exit Do
            Case Index > m_length
                m_parserrors = m_parserrors & "Falta '}': " & Right$(str, 20) & vbCrLf
                Exit Do
        End Select
        ' add key/value pair
        sKey = parseKey(Index)
        Call parseObject.Add(sKey, parseValue(str, Index))
        If Err.Number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & sKey & vbCrLf
            Exit Do
        End If
    Loop
    Exit Function
parseObject_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseObject", Erl)
End Function

Private Function parseArray(ByRef str As String, ByRef Index As Long) As Collection
    On Error GoTo parseArray_Err
    Dim charint As Integer
    Set parseArray = New Collection
    Call skipChar(Index)
    If mid$(str, Index, 1) <> "[" Then
        m_parserrors = m_parserrors & "Array invalido en la posicion " & Index & " : " + mid$(str, Index, 20) & vbCrLf
        Exit Function
    End If
    Index = Index + 1
    Do
        Call skipChar(Index)
        charint = m_str(Index)
        If charint = A_SQUARE_BRACKET_CLOSE Then
            Index = Index + 1
            Exit Do
        ElseIf charint = A_COMMA Then
            Index = Index + 1
            Call skipChar(Index)
        ElseIf Index > m_length Then
            m_parserrors = m_parserrors & "Falta ']': " & Right$(str, 20) & vbCrLf
            Exit Do
        End If
        'add value
        Call parseArray.Add(parseValue(str, Index))
        If Err.Number <> 0 Then
            m_parserrors = m_parserrors & Err.Description & ": " & mid$(str, Index, 20) & vbCrLf
            Exit Do
        End If
    Loop
    Exit Function
parseArray_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseArray", Erl)
End Function

Private Function parseValue(ByRef str As String, ByRef Index As Long)
    On Error GoTo parseValue_Err
    Call skipChar(Index)
    Select Case m_str(Index)
        Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
            parseValue = parseString(str, Index)
            Exit Function
        Case A_SQUARE_BRACKET_OPEN
            Set parseValue = parseArray(str, Index)
            Exit Function
        Case A_t, A_f
            parseValue = parseBoolean(str, Index)
            Exit Function
        Case A_n
            parseValue = parseNull(str, Index)
            Exit Function
        Case A_CURLY_BRACKET_OPEN
            Set parseValue = parseObject(str, Index)
            Exit Function
        Case Else
            parseValue = parseNumber(str, Index)
            Exit Function
    End Select
    Exit Function
parseValue_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseValue", Erl)
End Function

Private Function parseString(ByRef str As String, ByRef Index As Long) As String
    On Error GoTo parseString_Err
    Dim quoteint As Integer
    Dim charint  As Integer
    Dim Code     As String
    Call skipChar(Index)
    quoteint = m_str(Index)
    Index = Index + 1
    Do While Index > 0 And Index <= m_length
        charint = m_str(Index)
        Select Case charint
            Case A_BACKSLASH
                Index = Index + 1
                charint = m_str(Index)
                Select Case charint
                    Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
                        parseString = parseString & ChrW$(charint)
                        Index = Index + 1
                    Case A_b
                        parseString = parseString & vbBack
                        Index = Index + 1
                    Case A_f
                        parseString = parseString & vbFormFeed
                        Index = Index + 1
                    Case A_n
                        parseString = parseString & vbLf
                        Index = Index + 1
                    Case A_r
                        parseString = parseString & vbCr
                        Index = Index + 1
                    Case A_t
                        parseString = parseString & vbTab
                        Index = Index + 1
                    Case A_u
                        Index = Index + 1
                        Code = mid$(str, Index, 4)
                        parseString = parseString & ChrW$(val("&h" + Code))
                        Index = Index + 4
                End Select
            Case quoteint
                Index = Index + 1
                Exit Function
            Case Else
                parseString = parseString & ChrW$(charint)
                Index = Index + 1
        End Select
    Loop
    Exit Function
parseString_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseString", Erl)
End Function

Private Function parseNumber(ByRef str As String, ByRef Index As Long)
    On Error GoTo parseNumber_Err
    Dim value As String
    Dim Char  As String
    Call skipChar(Index)
    Do While Index > 0 And Index <= m_length
        Char = mid$(str, Index, 1)
        If InStr("+-0123456789.eE", Char) Then
            value = value & Char
            Index = Index + 1
        Else
            'check what is the grouping seperator
            If Not m_decSep = "." Then
                value = Replace(value, ".", m_decSep)
            End If
            If m_groupSep = "." Then
                value = Replace(value, ".", m_decSep)
            End If
            parseNumber = CDec(value)
            Exit Function
        End If
    Loop
    Exit Function
parseNumber_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseNumber", Erl)
End Function

Private Function parseBoolean(ByRef str As String, ByRef Index As Long) As Boolean
    On Error GoTo parseBoolean_Err
    Call skipChar(Index)
    If mid$(str, Index, 4) = "true" Then
        parseBoolean = True
        Index = Index + 4
    ElseIf mid$(str, Index, 5) = "false" Then
        parseBoolean = False
        Index = Index + 5
    Else
        m_parserrors = m_parserrors & "Boolean invalido en la posicion " & Index & " : " & mid$(str, Index) & vbCrLf
    End If
    Exit Function
parseBoolean_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseBoolean", Erl)
End Function

Private Function parseNull(ByRef str As String, ByRef Index As Long)
    On Error GoTo parseNull_Err
    Call skipChar(Index)
    If mid$(str, Index, 4) = "null" Then
        parseNull = Null
        Index = Index + 4
    Else
        m_parserrors = m_parserrors & "Valor nulo invalido en la posicion " & Index & " : " & mid$(str, Index) & vbCrLf
    End If
    Exit Function
parseNull_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseNull", Erl)
End Function

Private Function parseKey(ByRef Index As Long) As String
    On Error GoTo parseKey_Err
    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim charint As Integer
    Call skipChar(Index)
    Do While Index > 0 And Index <= m_length
        charint = m_str(Index)
        Select Case charint
            Case A_DOUBLE_QUOTE
                dquote = Not dquote
                Index = Index + 1
                If Not dquote Then
                    Call skipChar(Index)
                    If m_str(Index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & Index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
            Case A_SINGLE_QUOTE
                squote = Not squote
                Index = Index + 1
                If Not squote Then
                    Call skipChar(Index)
                    If m_str(Index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & Index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
            Case A_COLON
                Index = Index + 1
                If Not dquote And Not squote Then
                    Exit Do
                Else
                    parseKey = parseKey & ChrW$(charint)
                End If
            Case Else
                If A_VBCRLF = charint Then
                ElseIf A_VBCR = charint Then
                ElseIf A_VBLF = charint Then
                ElseIf A_VBTAB = charint Then
                ElseIf A_SPACE = charint Then
                Else
                    parseKey = parseKey & ChrW$(charint)
                End If
                Index = Index + 1
        End Select
    Loop
    Exit Function
parseKey_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.parseKey", Erl)
End Function

Private Sub skipChar(ByRef Index As Long)
    On Error GoTo skipChar_Err
    Dim bComment      As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment  As Boolean
    Do While Index > 0 And Index <= m_length
        Select Case m_str(Index)
            Case A_VBCR, A_VBLF
                If Not bLongComment Then
                    bStartComment = False
                    bComment = False
                End If
            Case A_VBTAB, A_SPACE, A_BRACKET_OPEN, A_BRACKET_CLOSE
                'do nothing
            Case A_FORWARDSLASH
                If Not bLongComment Then
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                    Else
                        bStartComment = True
                        bComment = False
                        bLongComment = False
                    End If
                Else
                    If bStartComment Then
                        bLongComment = False
                        bStartComment = False
                        bComment = False
                    End If
                End If
            Case A_ASTERIX
                If bStartComment Then
                    bStartComment = False
                    bComment = True
                    bLongComment = True
                Else
                    bStartComment = True
                End If
            Case Else
                If Not bComment Then
                    Exit Do
                End If
        End Select
        Index = Index + 1
    Loop
    Exit Sub
skipChar_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.skipChar", Erl)
End Sub

Public Function GetRegionalSettings(ByVal regionalsetting As Long) As String
    ' Devuelve la configuracion regional del sistema
    On Error GoTo ErrorHandler
    Dim Locale      As Long
    Dim Symbol      As String
    Dim iRet1       As Long
    Dim iRet2       As Long
    Dim lpLCDataVar As String
    Dim pos         As Integer
    Locale = GetUserDefaultLCID()
    iRet1 = GetLocaleInfo(Locale, regionalsetting, lpLCDataVar, 0)
    Symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, regionalsetting, Symbol, iRet1)
    pos = InStr(Symbol, Chr$(0))
    If pos > 0 Then
        Symbol = Left$(Symbol, pos - 1)
    End If
ErrorHandler:
    GetRegionalSettings = Symbol
    Select Case Err.Number
        Case 0
        Case Else
            Call Err.raise(123, "GetRegionalSetting", "GetRegionalSetting: " & regionalsetting)
    End Select
End Function

'********************************************************************************************************
'                   FUNCIONES MISCELANEAS DE LA ANTERIOR VERSION DEL MODULO
'********************************************************************************************************
Private Function Encode(str) As String
    On Error GoTo Encode_Err
    Dim SB  As New cStringBuilder
    Dim i   As Long
    Dim j   As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c   As String
    Dim p   As Boolean
    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To LenB(str)
        p = True
        c = mid$(str, i, 1)
        For j = 0 To 7
            If c = Chr$(aL1(j)) Then
                Call SB.Append("\" & Chr$(aL2(j)))
                p = False
                Exit For
            End If
        Next
        If p Then
            Dim a As Integer
            a = AscW(c)
            If a > 31 And a < 127 Then
                Call SB.Append(c)
            ElseIf a > -1 Or a < 65535 Then
                Call SB.Append("\u" & String$(4 - LenB(Hex$(a)), "0") & Hex$(a))
            End If
        End If
    Next
    Encode = SB.ToString
    Set SB = Nothing
    Exit Function
Encode_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.Encode", Erl)
End Function

Public Function StringToJSON(st As String) As String
    On Error GoTo StringToJSON_Err
    Const FIELD_SEP = "~"
    Const RECORD_SEP = "|"
    Dim sFlds     As String
    Dim sRecs     As New cStringBuilder
    Dim lRecCnt   As Long
    Dim lFld      As Long
    Dim fld       As Variant
    Dim rows      As Variant
    Dim Lower_fld As Long, Upper_fld As Long
    lRecCnt = 0
    If LenB(st) = 0 Then
        StringToJSON = "null"
    Else
        rows = Split(st, RECORD_SEP)
        For lRecCnt = LBound(rows) To UBound(rows)
            sFlds = vbNullString
            fld = Split(rows(lRecCnt), FIELD_SEP)
            For lFld = LBound(fld) To UBound(fld) Step 2
                sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
            Next 'fld
            Call sRecs.Append(IIf((Trim$(sRecs.ToString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}")
        Next 'rec
        StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.ToString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")
    End If
    Exit Function
StringToJSON_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.StringToJSON", Erl)
End Function

Public Function RStoJSON(RS As ADODB.Recordset) As String
    On Error GoTo ErrHandler
    Dim sFlds   As String
    Dim sRecs   As New cStringBuilder
    Dim lRecCnt As Long
    Dim fld     As ADODB.Field
    lRecCnt = 0
    If RS.State = adStateClosed Then
        RStoJSON = "null"
    Else
        If RS.EOF Or RS.BOF Then
            RStoJSON = "null"
        Else
            Do While Not RS.EOF And Not RS.BOF
                lRecCnt = lRecCnt + 1
                sFlds = vbNullString
                For Each fld In RS.Fields
                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.name & """:""" & toUnicode(fld.value & "") & """")
                Next 'fld
                Call sRecs.Append(IIf((Trim$(sRecs.ToString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}")
                Call RS.MoveNext
            Loop
            RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.ToString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")
        End If
    End If
    Exit Function
ErrHandler:
End Function

Public Function toUnicode(str As String) As String
    On Error GoTo toUnicode_Err
    Dim x        As Long
    Dim uStr     As New cStringBuilder
    Dim uChrCode As Integer
    For x = 1 To LenB(str)
        uChrCode = Asc(mid$(str, x, 1))
        Select Case uChrCode
            Case 8:   ' backspace
                Call uStr.Append("\b")
            Case 9: ' tab
                Call uStr.Append("\t")
            Case 10:  ' line feed
                Call uStr.Append("\n")
            Case 12:  ' formfeed
                Call uStr.Append("\f")
            Case 13: ' carriage return
                Call uStr.Append("\r")
            Case 34: ' quote
                Call uStr.Append("\""")
            Case 39:  ' apostrophe
                Call uStr.Append("\'")
            Case 92: ' backslash
                Call uStr.Append("\\")
            Case 123, 125:  ' "{" and "}"
                Call uStr.Append("\u" & Right$("0000" & Hex$(uChrCode), 4))
            Case Is < 32, Is > 127: ' non-ascii characters
                Call uStr.Append("\u" & Right$("0000" & Hex$(uChrCode), 4))
            Case Else
                Call uStr.Append(Chr$(uChrCode))
        End Select
    Next
    toUnicode = uStr.ToString
    Exit Function
    Exit Function
toUnicode_Err:
    Call TraceError(Err.Number, Err.Description, "mod_JSON.toUnicode", Erl)
End Function
