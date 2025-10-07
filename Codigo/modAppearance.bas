Attribute VB_Name = "modAppearance"
Option Explicit

Private Type tRange
    Lo As Integer
    Hi As Integer
End Type

Private Sub EnsureRandom()
    Static seeded As Boolean
    If Not seeded Then
        Randomize timer
        seeded = True
    End If
End Sub

Private Function PickFromRanges(ByRef rng() As tRange) As Integer
    ' Weighted uniformly over every valid head id across all ranges
    Dim i As Long, total As Long, size As Long, pick As Long
    EnsureRandom
    For i = LBound(rng) To UBound(rng)
        total = total + (rng(i).Hi - rng(i).Lo + 1)
    Next
    If total <= 0 Then
        PickFromRanges = 1
        Exit Function
    End If
    pick = Int(total * Rnd)  ' 0 .. total-1
    For i = LBound(rng) To UBound(rng)
        size = rng(i).Hi - rng(i).Lo + 1
        If pick < size Then
            PickFromRanges = rng(i).Lo + pick
            Exit Function
        End If
        pick = pick - size
    Next
    PickFromRanges = rng(LBound(rng)).Lo ' fallback
End Function

Public Function GetRandomHead(ByVal UserRaza As e_Raza, ByVal UserSexo As e_Genero) As Integer
    Dim ranges() As tRange
    Select Case UserSexo
        Case e_Genero.Hombre
            Select Case UserRaza
                Case e_Raza.Humano
                    ReDim ranges(0 To 1)
                    ranges(0).Lo = 1:   ranges(0).Hi = 41
                    ranges(1).Lo = 778: ranges(1).Hi = 791
                Case e_Raza.Elfo
                    ReDim ranges(0 To 1)
                    ranges(0).Lo = 101: ranges(0).Hi = 132
                    ranges(1).Lo = 531: ranges(1).Hi = 545
                Case e_Raza.Drow
                    ReDim ranges(0 To 1)
                    ranges(0).Lo = 200: ranges(0).Hi = 229
                    ranges(1).Lo = 792: ranges(1).Hi = 810
                Case e_Raza.Enano
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 300: ranges(0).Hi = 344
                Case e_Raza.Gnomo
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 400: ranges(0).Hi = 429
                Case e_Raza.Orco
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 500: ranges(0).Hi = 529
                Case Else
                    GetRandomHead = 1: Exit Function
            End Select

        Case e_Genero.Mujer
            Select Case UserRaza
                Case e_Raza.Humano
                    ReDim ranges(0 To 2)
                    ranges(0).Lo = 50:  ranges(0).Hi = 80
                    ranges(1).Lo = 187: ranges(1).Hi = 190
                    ranges(2).Lo = 230: ranges(2).Hi = 246
                Case e_Raza.Elfo
                    ReDim ranges(0 To 1)
                    ranges(0).Lo = 150: ranges(0).Hi = 179
                    ranges(1).Lo = 758: ranges(1).Hi = 777
                Case e_Raza.Drow
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 250: ranges(0).Hi = 279
                Case e_Raza.Enano
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 350: ranges(0).Hi = 379
                Case e_Raza.Gnomo
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 450: ranges(0).Hi = 479
                Case e_Raza.Orco
                    ReDim ranges(0 To 0)
                    ranges(0).Lo = 550: ranges(0).Hi = 579
                Case Else
                    GetRandomHead = 50: Exit Function
            End Select

        Case Else
            GetRandomHead = 1: Exit Function
    End Select

    GetRandomHead = PickFromRanges(ranges)
End Function

