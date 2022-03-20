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

Public public_key As String
Public private_key As String

Public encrypted_token As String
Public decrypted_token As String

Public character_name As String

Public Sub init()
    'We can mock the key value to test errors...
    private_key = PrivateKey
    character_name = "seneca"
    character_name = RandomName(16)
    'Hardcoded token for unit testing...

    decrypted_token = "G7H5wKOKZvebZxHtnkRtJNvL/AHWEw3dHCyBTzXVvdTe3bQAJHePsFfV/Ecgm9Wk"
    encrypted_token = AO20CryptoSysWrapper.ENCRYPT(private_key, decrypted_token)
    public_key = mid$(decrypted_token, 1, 16)
    
    'Add a fake token to be using when exercising the protocol for LoginNewChar
    Call AddTokenDatabase(encrypted_token, decrypted_token, "MORGOLOCK2002@YAHOO.COM.AR")
    
End Sub

Public Sub shutdown()
    Call UnitClient.Disconnect
End Sub

Sub test_make_user(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal y As Integer)
    UserList(UserIndex).Pos.map = map
    UserList(UserIndex).Pos.X = X
    UserList(UserIndex).Pos.y = y
    Call MakeUserChar(True, 17, UserIndex, map, X, y, 1)
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
    Dim i As Integer
    For i = 1 To 100
            Debug.Assert Porcentaje(100#, i) = i
    Next i
    For i = 1 To 1000
            Debug.Assert Porcentaje(1000#, i) = i * 10
    Next i
    Debug.Print "Porcentaje took " & sw.ElapsedMilliseconds; " ms"
    test_percentage = True
End Function

Function test_distance() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    Debug.Assert Distance(0, 0, 0, 0) = 0
    Dim i As Integer
    For i = 1 To 100
            Debug.Assert Distance(i, 0, 0, 0) = i
    Next i
    For i = 1 To 1000
           Debug.Assert Distance(i, 0, -i, 0) = i + i
    Next i
    Debug.Print "distace took " & sw.ElapsedMilliseconds; " ms"
    test_distance = True
End Function


Function test_random_number() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Debug.Assert RandomNumber(0, 0) = 0
    Debug.Assert RandomNumber(-1, -1) = -1
    Debug.Assert RandomNumber(1, 1) = 1
    Dim i As Integer
    Dim n As Integer
    For i = 1 To 1000
          n = RandomNumber(0, i)
          Debug.Assert n >= 0 And n <= i
    Next i
    For i = 1 To 1000
          n = RandomNumber(-i, 0)
          Debug.Assert n >= -i And n <= 0
    Next i
    
    Debug.Print "random_bumber took " & sw.ElapsedMilliseconds; " ms"
    test_random_number = True
End Function


Function test_maths() As Boolean
    test_maths = test_percentage() And test_random_number() And test_distance()
End Function

Function test_make_user_char() As Boolean
    'Create first User
    Call test_make_user(1, 1, 54, 51)
    Debug.Assert (MapData(1, 54, 51).UserIndex = 1)
    Debug.Assert (UserList(1).Char.CharIndex <> 0)
    'Delete first user
    Call EraseUserChar(1, False, False)
    Debug.Assert (MapData(1, 54, 55).UserIndex = 0)
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
    Debug.Assert (MapData(1, 54, 56).UserIndex = 2)
    Debug.Assert (UserList(2).Char.CharIndex <> 0)
    
    Call test_make_user(1, 1, 50, 46)
    Debug.Assert (MapData(1, 50, 46).UserIndex = 1)
    Debug.Assert (UserList(1).Char.CharIndex <> 0)
    Debug.Assert (UserList(2).Char.CharIndex <> UserList(1).Char.CharIndex)
    
    'Delete user 2
    Call EraseUserChar(2, False, False)
    Debug.Assert (MapData(1, 54, 56).UserIndex = 0)
    Debug.Assert (UserList(2).Char.CharIndex = 0)
    'Create user 2 again
    Call test_make_user(2, 1, 54, 56)
    Debug.Assert (MapData(1, 54, 56).UserIndex = 2)
    Debug.Assert (UserList(2).Char.CharIndex <> 0)
    
    For i = 1 To UBound(UserList)
        If UserList(i).Char.CharIndex <> 0 Then
            Call EraseUserChar(i, False, True)
        End If
    Next i
    
    Call test_make_user(1, 1, 64, 66)
    Debug.Assert (MapData(1, 64, 66).UserIndex = 1)
    Debug.Assert (UserList(1).Char.CharIndex <> 0)
    Debug.Assert (UserList(1).Char.CharIndex = 1)
    
    
    Call test_make_user(1, 1, 68, 66)
    Debug.Assert (MapData(1, 68, 66).UserIndex = 1)
    Debug.Assert (UserList(1).Char.CharIndex <> 0)
    test_make_user_char = True
End Function

Function test_suite() As Boolean
    Dim result As Boolean
    result = test_make_user_char()
    result = result And test_maths()
    test_suite = result
End Function

