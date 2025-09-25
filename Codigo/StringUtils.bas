Attribute VB_Name = "StringUtils"
Function ValidWordsDescription(ByVal cad As String) As Boolean
    On Error GoTo ValidWordsDescription_Err
    Dim i As Integer
    ' Convertimos todo a minúsculas
    cad = LCase$(cad)
    ' Agregamos espacios al inicio y final para asegurar coincidencias exactas de palabras/frases
    cad = " " & NormalizeText(cad) & " "
    ' Verificamos si alguna palabra/frase prohibida está contenida en la descripción
    For i = LBound(BlockedWordsDescription) To UBound(BlockedWordsDescription)
        If InStr(1, cad, " " & BlockedWordsDescription(i) & " ", vbTextCompare) > 0 Then
            ValidWordsDescription = False
            Exit Function
        End If
    Next i
    ValidWordsDescription = True
    Exit Function
ValidWordsDescription_Err:
    Call TraceError(Err.Number, Err.Description, "StringUtils.ValidWordsDescription", Erl)
End Function

Private Function NormalizeText(ByVal cad As String) As String
    On Error GoTo NormalizeText_Err
    ' Esta función normaliza una cadena para facilitar la detección de palabras/frases prohibidas
    Dim PunctuationMarks As String
    Dim i                As Integer
    ' Lista de signos que queremos reemplazar por espacios
    PunctuationMarks = ".,;:!?()[]<>-/_\"
    ' Convertimos todo el texto a minúsculas para evitar diferencias por mayúsculas
    cad = LCase$(cad)
    ' Recorremos cada signo y lo reemplazamos por un espacio
    For i = 1 To Len(PunctuationMarks)
        cad = Replace(cad, mid$(PunctuationMarks, i, 1), " ")
    Next i
    ' Reemplazamos espacios dobles (o múltiples) por espacios simples
    Do While InStr(cad, "  ") > 0
        cad = Replace(cad, "  ", " ")
    Loop
    ' Quitamos espacios al inicio y final de la cadena
    NormalizeText = Trim$(cad)
    Exit Function
NormalizeText_Err:
    Call TraceError(Err.Number, Err.Description, "StringUtils.NormalizeText", Erl)
End Function

Function ValidDescription(ByVal cad As String) As Boolean
    On Error GoTo ValidDescription_Err
    Dim car As Byte
    Dim i   As Integer
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If car < 32 Or car >= 126 Then
            ValidDescription = False
            Exit Function
        End If
    Next i
    ValidDescription = True
    Exit Function
ValidDescription_Err:
    Call TraceError(Err.Number, Err.Description, "StringUtils.ValidDescription", Erl)
End Function
