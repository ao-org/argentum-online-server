Attribute VB_Name = "UnitTesting"
Option Explicit


Sub TestMakeUserChar()
Dim u_map, u_userindex, u_charindex, u_userindex2, u_posx, u_posy
u_map = 1
u_posx = 54
u_posy = 51
u_userindex = 1
u_userindex2 = 2

UserList(u_userindex).Pos.Map = u_map
UserList(u_userindex).Pos.X = u_posx
UserList(u_userindex).Pos.Y = u_posy

Call MakeUserChar(True, 17, u_userindex, u_map, u_posx, u_posy, 1)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = u_userindex)
u_charindex = UserList(u_userindex).Char.CharIndex
Debug.Assert (UserList(u_userindex).Char.CharIndex = u_charindex)
Call EraseUserChar(u_userindex, False, False)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = 0)

'Delete all NPCs
Dim i
For i = 1 To UBound(NpcList)
        If NpcList(i).Char.CharIndex <> 0 Then
            Call EraseNPCChar(1)
        End If
Next i

Call MakeUserChar(True, 17, u_userindex2, u_map, u_posx, u_posy, 1)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = u_userindex2)
Call MakeUserChar(True, 17, u_userindex, u_map, u_posx, u_posy, 1)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = u_userindex)
Debug.Assert (UserList(u_userindex2).Char.CharIndex <> UserList(u_userindex).Char.CharIndex)
Call EraseUserChar(u_userindex2, False, True)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = 0)
Call MakeUserChar(True, 17, u_userindex2, u_map, u_posx, u_posy, 1)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = u_userindex2)
Debug.Assert (UserList(u_userindex2).Char.CharIndex <> 1)

For i = 1 To UBound(UserList)
    If UserList(i).Char.CharIndex <> 0 Then
        Call EraseUserChar(UserList(i).Char.CharIndex, False, True)
    End If
Next i
Call MakeUserChar(True, 17, u_userindex, u_map, u_posx, u_posy, 1)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = u_userindex)
u_charindex = UserList(u_userindex).Char.CharIndex
Debug.Assert (UserList(u_userindex).Char.CharIndex = u_charindex)
Debug.Assert (UserList(u_userindex).Char.CharIndex = 1)

End Sub

Function TestSuite()

TestMakeUserChar

TestSuite = True
End Function

