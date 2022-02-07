Attribute VB_Name = "UnitTesting"
Option Explicit


Sub test_make_user(ByVal userindex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer)
UserList(userindex).Pos.map = map
UserList(userindex).Pos.x = x
UserList(userindex).Pos.y = y
Call MakeUserChar(True, 17, userindex, map, x, y, 1)
End Sub

Function TestMakeUserChar() As Boolean

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
TestMakeUserChar = True
End Function

Function TestSuite() As Boolean


TestSuite = TestMakeUserChar()

End Function

