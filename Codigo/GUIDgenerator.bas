Attribute VB_Name = "GUIDgenerator"
Option Explicit

Private MyGUID(35) As Byte
Private DashNum As Byte
Private FourNum As Byte

Public Sub GenInit()
    Randomize
    FourNum = Asc("4")
    DashNum = Asc("-")
End Sub

Public Function GenGUID() As String
    Dim i As Integer

    For i = 0 To 7
    MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i
    
    For i = 9 To 12
    MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i
    
    For i = 15 To 17
    MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i
    
    For i = 19 To 22
    MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i
    
    For i = 24 To 35
    MyGUID(i) = Asc(Hex(Int(16 * Rnd())))
    Next i
    
    MyGUID(8) = DashNum
    MyGUID(13) = DashNum
    MyGUID(14) = FourNum
    MyGUID(18) = DashNum
    MyGUID(23) = DashNum
    
    GenGUID = "{" & StrConv(MyGUID, vbUnicode) & "}"
End Function
