Attribute VB_Name = "ModEncrypt"
Public Function SEncriptar(ByVal Cadena As String) As String
' GSZ-AO - Encripta una cadena de texto
    Dim i As Long, RandomNum As Integer
    
    RandomNum = 99 * Rnd
    If RandomNum < 10 Then RandomNum = 10
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i
    SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
    'DoEvents (WyroX: WTF?)

End Function

Public Function SDesencriptar(ByVal Cadena As String) As String
' GSZ-AO - Desencripta una cadena de texto
    Dim i As Long, NumDesencriptar As String
    
    NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
    Cadena = (Left$(Cadena, Len(Cadena) - 2))
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
    Next i
    SDesencriptar = Cadena
    'DoEvents (WyroX: WTF?)

End Function
' GSZAO - Encriptación basica y rapida para Strings
Public Function RndCrypt(ByVal str As String, ByVal Password As String) As String
    '  Made by Michael Ciurescu
    ' (CVMichael from vbforums.com)
    '  Original thread: http://www.vbforums.com/showthread.php?t=231798
    Dim SK As Long, K As Long

    Rnd -1
    Randomize Len(Password)

    For K = 1 To Len(Password)
        SK = SK + (((K Mod 256) _
        Xor Asc(mid$(Password, K, 1))) _
        Xor Fix(256 * Rnd))
    Next K

    Rnd -1
    Randomize SK
    
    For K = 1 To Len(str)
        Mid$(str, K, 1) = Chr(Fix(256 * Rnd) _
        Xor Asc(mid$(str, K, 1)))
    Next K
    
    RndCrypt = str
End Function

