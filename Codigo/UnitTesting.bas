Attribute VB_Name = "UnitTesting"
Option Explicit


Sub TestMakeUserChar()
Dim u_map, u_userindex, u_posx, u_posy
u_map = 1
u_posx = 54
u_posy = 51
u_userindex = 1
Call MakeUserChar(True, 17, u_userindex, u_map, u_posx, u_posy, 1)
Debug.Assert (MapData(u_map, u_posx, u_posy).UserIndex = u_userindex)
End Sub

Function TestSuite()

TestMakeUserChar

TestSuite = True
End Function

