Attribute VB_Name = "UnitTesting"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit




Sub test_make_user(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
UserList(userindex).Pos.map = map
UserList(userindex).Pos.x = x
UserList(userindex).Pos.y = y
Call MakeUserChar(True, 17, userindex, map, x, y, 1)
End Sub

Function test_percentage() As Boolean
Dim sw As Instruments
Set sw = New Instruments
sw.start
Debug.Assert (Porcentaje(100#, 1#) = 1)
Debug.Assert (Porcentaje(100#, 2#) = 2)
Debug.Assert (Porcentaje(100#, 5#) = 5)
Debug.Assert (Porcentaje(100#, 10#) = 10)
Debug.Assert (Porcentaje(100#, 100#) = 100)
Debug.Print "Porcentaje took " & sw.ElapsedMilliseconds; " ms"
test_percentage = True
End Function

Function test_maths() As Boolean
test_maths = test_percentage()
End Function

Function test_make_user_char() As Boolean

'Create first User
Call test_make_user(1, 1, 54, 51)
Debug.Assert (MapData(1, 54, 51).userindex = 1)
Debug.Assert (UserList(1).Char.CharIndex <> 0)
'Delete first user
Call EraseUserChar(1, False, False)
Debug.Assert (MapData(1, 54, 55).userindex = 0)
Debug.Assert (UserList(1).Char.CharIndex = 0)
'Delete all NPCs5
Dim i
For i = 1 To UBound(NpcList)
        If NpcList(i).Char.CharIndex <> 0 Then
            Call EraseNPCChar(1)
        End If
Next i

'Create two users on the same map pos
Call test_make_user(2, 1, 54, 56)
Debug.Assert (MapData(1, 54, 56).userindex = 2)
Debug.Assert (UserList(2).Char.CharIndex <> 0)

Call test_make_user(1, 1, 50, 46)
Debug.Assert (MapData(1, 50, 46).userindex = 1)
Debug.Assert (UserList(1).Char.CharIndex <> 0)
Debug.Assert (UserList(2).Char.CharIndex <> UserList(1).Char.CharIndex)

'Delete user 2
Call EraseUserChar(2, False, False)
Debug.Assert (MapData(1, 54, 56).userindex = 0)
Debug.Assert (UserList(2).Char.CharIndex = 0)
'Create user 2 again
Call test_make_user(2, 1, 54, 56)
Debug.Assert (MapData(1, 54, 56).userindex = 2)
Debug.Assert (UserList(2).Char.CharIndex <> 0)

For i = 1 To UBound(UserList)
    If UserList(i).Char.CharIndex <> 0 Then
        Call EraseUserChar(i, False, True)
    End If
Next i

Call test_make_user(1, 1, 64, 66)
Debug.Assert (MapData(1, 64, 66).userindex = 1)
Debug.Assert (UserList(1).Char.CharIndex <> 0)
Debug.Assert (UserList(1).Char.CharIndex = 1)


Call test_make_user(1, 1, 68, 66)
Debug.Assert (MapData(1, 68, 66).userindex = 1)
Debug.Assert (UserList(1).Char.CharIndex <> 0)
test_make_user_char = True
End Function

Function test_suite() As Boolean

Dim result As Boolean

result = test_make_user_char()
result = result And test_maths()
test_suite = result
End Function

