Attribute VB_Name = "mod_JSON"

' VBJSONDeserializer is a VB6 adaptation of the VB-JSON project @
' Fuente: https://www.codeproject.com/Articles/720368/VB-JSON-Parser-Improved-Performance

' BSD Licensed

Option Explicit

' DECLARACIONES API
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

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
    
        
100     GetParserErrors = m_parserrors
        
        Exit Function

GetParserErrors_Err:
102     Call RegistrarError(Err.Number, Err.description, "mod_JSON.GetParserErrors", Erl)

        
End Function

Public Function parse(ByRef str As String) As Object
        
        On Error GoTo parse_Err
    
        

100     m_decSep = GetRegionalSettings(LOCALE_SDECIMAL)
102     m_groupSep = GetRegionalSettings(LOCALE_SGROUPING)

        Dim index As Long
104         index = 1

106     Call GenerateStringArray(str)

108     m_parserrors = vbNullString

    

110     Call skipChar(index)

112     Select Case m_str(index)

            Case A_SQUARE_BRACKET_OPEN
114             Set parse = parseArray(str, index)

116         Case A_CURLY_BRACKET_OPEN
118             Set parse = parseObject(str, index)

120         Case Else
122             m_parserrors = "JSON Invalido"

        End Select

        'clean array
124     ReDim m_str(1)

        
        Exit Function

parse_Err:
126     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parse", Erl)

        
End Function

Private Sub GenerateStringArray(ByRef str As String)
        
        On Error GoTo GenerateStringArray_Err
    
        

        Dim i As Long

100     m_length = Len(str)
102     ReDim m_str(1 To m_length)

104     For i = 1 To m_length
106         m_str(i) = AscW(mid$(str, i, 1))
108     Next i

        
        Exit Sub

GenerateStringArray_Err:
110     Call RegistrarError(Err.Number, Err.description, "mod_JSON.GenerateStringArray", Erl)

        
End Sub

Private Function parseObject(ByRef str As String, ByRef index As Long) As Dictionary
        
        On Error GoTo parseObject_Err
    
        

100     Set parseObject = New Dictionary

        Dim sKey    As String
        Dim charint As Integer

102     Call skipChar(index)

104     If m_str(index) <> A_CURLY_BRACKET_OPEN Then
106         m_parserrors = m_parserrors & "Objeto invalido en la posicion " & index & " : " & mid$(str, index) & vbCrLf
            Exit Function

        End If

108     index = index + 1

        Do
110         Call skipChar(index)
    
112         charint = m_str(index)
        
114         Select Case charint
        
                Case A_COMMA
116                 index = index + 1
118                 Call skipChar(index)
            
120             Case A_CURLY_BRACKET_CLOSE
122                 index = index + 1
                    Exit Do
                
124             Case index > m_length
126                 m_parserrors = m_parserrors & "Falta '}': " & Right$(str, 20) & vbCrLf
                    Exit Do
                
            End Select

            ' add key/value pair
128         sKey = parseKey(index)

        

130         Call parseObject.Add(sKey, parseValue(str, index))

132         If Err.Number <> 0 Then
134             m_parserrors = m_parserrors & Err.description & ": " & sKey & vbCrLf
                Exit Do
            End If

        Loop

        
        Exit Function

parseObject_Err:
136     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseObject", Erl)

        
End Function

Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection
        
        On Error GoTo parseArray_Err
    
        

        Dim charint As Integer

100     Set parseArray = New Collection

102     Call skipChar(index)

104     If mid$(str, index, 1) <> "[" Then
106         m_parserrors = m_parserrors & "Array invalido en la posicion " & index & " : " + mid$(str, index, 20) & vbCrLf
            Exit Function
        End If
   
108     index = index + 1

        Do
110         Call skipChar(index)
    
112         charint = m_str(index)
    
114         If charint = A_SQUARE_BRACKET_CLOSE Then
116             index = index + 1
                Exit Do
118         ElseIf charint = A_COMMA Then
120             index = index + 1
122             Call skipChar(index)
124         ElseIf index > m_length Then
126             m_parserrors = m_parserrors & "Falta ']': " & Right$(str, 20) & vbCrLf
                Exit Do
            End If
    
            'add value
        

128         Call parseArray.Add(parseValue(str, index))

130         If Err.Number <> 0 Then
132             m_parserrors = m_parserrors & Err.description & ": " & mid$(str, index, 20) & vbCrLf
                Exit Do

            End If

        Loop

        
        Exit Function

parseArray_Err:
134     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseArray", Erl)

        
End Function

Private Function parseValue(ByRef str As String, ByRef index As Long)
        
        On Error GoTo parseValue_Err
    
        

100     Call skipChar(index)

102     Select Case m_str(index)

            Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
104             parseValue = parseString(str, index)
                Exit Function

106         Case A_SQUARE_BRACKET_OPEN
108             Set parseValue = parseArray(str, index)
                Exit Function

110         Case A_t, A_f
112             parseValue = parseBoolean(str, index)
                Exit Function

114         Case A_n
116             parseValue = parseNull(str, index)
                Exit Function

118         Case A_CURLY_BRACKET_OPEN
120             Set parseValue = parseObject(str, index)
                Exit Function

122         Case Else
124             parseValue = parseNumber(str, index)
                Exit Function

        End Select

        
        Exit Function

parseValue_Err:
126     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseValue", Erl)

        
End Function

Private Function parseString(ByRef str As String, ByRef index As Long) As String
        
        On Error GoTo parseString_Err
    
        

        Dim quoteint As Integer
        Dim charint  As Integer
        Dim code     As String
   
100     Call skipChar(index)
   
102     quoteint = m_str(index)
   
104     index = index + 1
   
106     Do While index > 0 And index <= m_length
   
108         charint = m_str(index)
      
110         Select Case charint

                Case A_BACKSLASH

112                 index = index + 1
114                 charint = m_str(index)

116                 Select Case charint

                        Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
118                         parseString = parseString & ChrW$(charint)
120                         index = index + 1

122                     Case A_b
124                         parseString = parseString & vbBack
126                         index = index + 1

128                     Case A_f
130                         parseString = parseString & vbFormFeed
132                         index = index + 1

134                     Case A_n
136                         parseString = parseString & vbLf
138                         index = index + 1

140                     Case A_r
142                         parseString = parseString & vbCr
144                         index = index + 1

146                     Case A_t
148                         parseString = parseString & vbTab
150                         index = index + 1

152                     Case A_u
154                         index = index + 1
156                         code = mid$(str, index, 4)

158                         parseString = parseString & ChrW$(val("&h" + code))
160                         index = index + 4

                    End Select

162             Case quoteint
164                 index = index + 1
                    Exit Function

166             Case Else
168                 parseString = parseString & ChrW$(charint)
170                 index = index + 1

            End Select

        Loop
   
        
        Exit Function

parseString_Err:
172     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseString", Erl)

        
End Function

Private Function parseNumber(ByRef str As String, ByRef index As Long)
        
        On Error GoTo parseNumber_Err
    
        

        Dim Value As String
        Dim Char  As String

100     Call skipChar(index)

102     Do While index > 0 And index <= m_length
104         Char = mid$(str, index, 1)

106         If InStr("+-0123456789.eE", Char) Then
108             Value = Value & Char
110             index = index + 1
            Else

                'check what is the grouping seperator
112             If Not m_decSep = "." Then
114                 Value = Replace(Value, ".", m_decSep)

                End If
     
116             If m_groupSep = "." Then
118                 Value = Replace(Value, ".", m_decSep)

                End If
     
120             parseNumber = CDec(Value)
                Exit Function

            End If

        Loop
   
        
        Exit Function

parseNumber_Err:
122     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseNumber", Erl)

        
End Function

Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean
        
        On Error GoTo parseBoolean_Err
    
        

100     Call skipChar(index)
   
102     If mid$(str, index, 4) = "true" Then
104         parseBoolean = True
106         index = index + 4
108     ElseIf mid$(str, index, 5) = "false" Then
110         parseBoolean = False
112         index = index + 5
        Else
114         m_parserrors = m_parserrors & "Boolean invalido en la posicion " & index & " : " & mid$(str, index) & vbCrLf

        End If

        
        Exit Function

parseBoolean_Err:
116     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseBoolean", Erl)

        
End Function

Private Function parseNull(ByRef str As String, ByRef index As Long)
        
        On Error GoTo parseNull_Err
    
        

100     Call skipChar(index)
   
102     If mid$(str, index, 4) = "null" Then
104         parseNull = Null
106         index = index + 4
        Else
108         m_parserrors = m_parserrors & "Valor nulo invalido en la posicion " & index & " : " & mid$(str, index) & vbCrLf

        End If

        
        Exit Function

parseNull_Err:
110     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseNull", Erl)

        
End Function

Private Function parseKey(ByRef index As Long) As String
        
        On Error GoTo parseKey_Err
    
        

        Dim dquote  As Boolean
        Dim squote  As Boolean
        Dim charint As Integer
   
100     Call skipChar(index)
   
102     Do While index > 0 And index <= m_length
    
104         charint = m_str(index)
        
106         Select Case charint

                Case A_DOUBLE_QUOTE
108                 dquote = Not dquote
110                 index = index + 1

112                 If Not dquote Then
            
114                     Call skipChar(index)
                
116                     If m_str(index) <> A_COLON Then
118                         m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & index & " : " & parseKey & vbCrLf
                            Exit Do

                        End If

                    End If

120             Case A_SINGLE_QUOTE
122                 squote = Not squote
124                 index = index + 1

126                 If Not squote Then
128                     Call skipChar(index)
                
130                     If m_str(index) <> A_COLON Then
132                         m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & index & " : " & parseKey & vbCrLf
                            Exit Do

                        End If
                
                    End If
        
134             Case A_COLON
136                 index = index + 1

138                 If Not dquote And Not squote Then
                        Exit Do
                    Else
140                     parseKey = parseKey & ChrW$(charint)

                    End If

142             Case Else
            
144                 If A_VBCRLF = charint Then
146                 ElseIf A_VBCR = charint Then
148                 ElseIf A_VBLF = charint Then
150                 ElseIf A_VBTAB = charint Then
152                 ElseIf A_SPACE = charint Then
                    Else
154                     parseKey = parseKey & ChrW$(charint)

                    End If

156                 index = index + 1

            End Select

        Loop

        
        Exit Function

parseKey_Err:
158     Call RegistrarError(Err.Number, Err.description, "mod_JSON.parseKey", Erl)

        
End Function

Private Sub skipChar(ByRef index As Long)
        
        On Error GoTo skipChar_Err
    
        

        Dim bComment      As Boolean
        Dim bStartComment As Boolean
        Dim bLongComment  As Boolean

100     Do While index > 0 And index <= m_length
    
102         Select Case m_str(index)

                Case A_VBCR, A_VBLF

104                 If Not bLongComment Then
106                     bStartComment = False
108                     bComment = False

                    End If
    
110             Case A_VBTAB, A_SPACE, A_BRACKET_OPEN, A_BRACKET_CLOSE
                    'do nothing
        
112             Case A_FORWARDSLASH

114                 If Not bLongComment Then
116                     If bStartComment Then
118                         bStartComment = False
120                         bComment = True
                        Else
122                         bStartComment = True
124                         bComment = False
126                         bLongComment = False

                        End If

                    Else

128                     If bStartComment Then
130                         bLongComment = False
132                         bStartComment = False
134                         bComment = False

                        End If

                    End If

136             Case A_ASTERIX

138                 If bStartComment Then
140                     bStartComment = False
142                     bComment = True
144                     bLongComment = True
                    Else
146                     bStartComment = True

                    End If

148             Case Else
        
150                 If Not bComment Then
                        Exit Do

                    End If

            End Select

152         index = index + 1
        Loop

        
        Exit Sub

skipChar_Err:
154     Call RegistrarError(Err.Number, Err.description, "mod_JSON.skipChar", Erl)

        
End Sub

Public Function GetRegionalSettings(ByVal regionalsetting As Long) As String
        ' Devuelve la configuracion regional del sistema

        On Error GoTo ErrorHandler

        Dim Locale      As Long
        Dim Symbol      As String
        Dim iRet1       As Long
        Dim iRet2       As Long
        Dim lpLCDataVar As String
        Dim Pos         As Integer
      
100     Locale = GetUserDefaultLCID()

102     iRet1 = GetLocaleInfo(Locale, regionalsetting, lpLCDataVar, 0)
104     Symbol = String$(iRet1, 0)
106     iRet2 = GetLocaleInfo(Locale, regionalsetting, Symbol, iRet1)
108     Pos = InStr(Symbol, Chr$(0))

110     If Pos > 0 Then
112         Symbol = Left$(Symbol, Pos - 1)
        End If
      
ErrorHandler:
114     GetRegionalSettings = Symbol

116     Select Case Err.Number

            Case 0

118         Case Else
120             Call Err.raise(123, "GetRegionalSetting", "GetRegionalSetting: " & regionalsetting)

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

100     aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
102     aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)

104     For i = 1 To LenB(str)
106         p = True
108         c = mid$(str, i, 1)

110         For j = 0 To 7

112             If c = Chr$(aL1(j)) Then
114                 Call SB.Append("\" & Chr$(aL2(j)))
116                 p = False
                    Exit For

                End If

            Next

118         If p Then

                Dim a As Integer
120                 a = AscW(c)

122             If a > 31 And a < 127 Then
124                 Call SB.Append(c)
126             ElseIf a > -1 Or a < 65535 Then
128                 Call SB.Append("\u" & String$(4 - LenB(hex$(a)), "0") & hex$(a))
                End If

            End If

        Next
   
130     Encode = SB.toString
    
132     Set SB = Nothing
   
        
        Exit Function

Encode_Err:
134     Call RegistrarError(Err.Number, Err.description, "mod_JSON.Encode", Erl)

        
End Function

Public Function StringToJSON(st As String) As String
        
        On Error GoTo StringToJSON_Err
    
        
   
        Const FIELD_SEP = "~"
        Const RECORD_SEP = "|"

        Dim sFlds   As String
        Dim sRecs   As New cStringBuilder
        Dim lRecCnt As Long
        Dim lFld    As Long
        Dim fld     As Variant
        Dim rows    As Variant
    
        Dim Lower_fld As Long, Upper_fld As Long

100     lRecCnt = 0

102     If LenB(st) = 0 Then
104         StringToJSON = "null"
        Else
106         rows = Split(st, RECORD_SEP)
        
108         For lRecCnt = LBound(rows) To UBound(rows)
110             sFlds = vbNullString
112             fld = Split(rows(lRecCnt), FIELD_SEP)

114             For lFld = LBound(fld) To UBound(fld) Step 2
116                 sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
                Next 'fld

118             Call sRecs.Append(IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}")
            Next 'rec

120         StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

        End If
        
        Exit Function

StringToJSON_Err:
122     Call RegistrarError(Err.Number, Err.description, "mod_JSON.StringToJSON", Erl)

        
End Function

Public Function RStoJSON(rs As ADODB.Recordset) As String

        On Error GoTo ErrHandler

        Dim sFlds   As String
        Dim sRecs   As New cStringBuilder
        Dim lRecCnt As Long
        Dim fld     As ADODB.Field

100     lRecCnt = 0

102     If rs.State = adStateClosed Then
104         RStoJSON = "null"
        Else

106         If rs.EOF Or rs.BOF Then
108             RStoJSON = "null"
            
            Else

110             Do While Not rs.EOF And Not rs.BOF
112                 lRecCnt = lRecCnt + 1
114                 sFlds = vbNullString

116                 For Each fld In rs.Fields
118                     sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.name & """:""" & toUnicode(fld.Value & "") & """")
                    Next 'fld

120                 Call sRecs.Append(IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}")
122                 Call rs.MoveNext
                Loop
            
124             RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")

            End If

        End If

        Exit Function
ErrHandler:

End Function

Public Function toUnicode(str As String) As String
        
        On Error GoTo toUnicode_Err
    
        

        Dim X        As Long
        Dim uStr     As New cStringBuilder
        Dim uChrCode As Integer

100     For X = 1 To LenB(str)
102         uChrCode = Asc(mid$(str, X, 1))

104         Select Case uChrCode

                Case 8:   ' backspace
106                 Call uStr.Append("\b")

108             Case 9: ' tab
110                 Call uStr.Append("\t")

112             Case 10:  ' line feed
114                 Call uStr.Append("\n")

116             Case 12:  ' formfeed
118                 Call uStr.Append("\f")

120             Case 13: ' carriage return
122                 Call uStr.Append("\r")

124             Case 34: ' quote
126                 Call uStr.Append("\""")

128             Case 39:  ' apostrophe
130                 Call uStr.Append("\'")

132             Case 92: ' backslash
134                 Call uStr.Append("\\")

136             Case 123, 125:  ' "{" and "}"
138                 Call uStr.Append("\u" & Right$("0000" & hex$(uChrCode), 4))

140             Case Is < 32, Is > 127: ' non-ascii characters
142                 Call uStr.Append("\u" & Right$("0000" & hex$(uChrCode), 4))

144             Case Else
146                 Call uStr.Append(Chr$(uChrCode))

            End Select

        Next
    
148     toUnicode = uStr.toString
    
        Exit Function

        
        Exit Function

toUnicode_Err:
150     Call RegistrarError(Err.Number, Err.description, "mod_JSON.toUnicode", Erl)

        
End Function
