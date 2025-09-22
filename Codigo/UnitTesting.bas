Attribute VB_Name = "UnitTesting"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
#If UNIT_TEST = 1 Then

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

Function test_npc_pathfinding_attackable_state() As Boolean
    Dim npcIndex As Integer
    Dim attackCheck As t_AttackInteractionResult
    npcIndex = 1

    Call ResetNpcMainInfo(npcIndex)
    Call ResetNpcFlags(npcIndex)
    Call ResetNpcCounters(npcIndex)
    ReDim NpcList(npcIndex).pathFindingInfo.Path(1 To MAX_PATH_LENGTH)

    With NpcList(npcIndex)
        .Attackable = 1
        .Hostile = 1
        .Pos.Map = 1
        .Pos.X = 10
        .Pos.Y = 10
        .Orig = .Pos
        .pathFindingInfo.RangoVision = 1
        .pathFindingInfo.OriginalVision = 1
        .pathFindingInfo.PathLength = 0
        .flags.Faccion = e_Facciones.Ciudadano
        .Humanoide = False
    End With

    NumMaps = 1
    MinXBorder = XMinMapSize
    MaxXBorder = XMaxMapSize
    MinYBorder = YMinMapSize
    MaxYBorder = YMaxMapSize
    ReDim MapData(1 To 1, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)
    ReDim MapInfo(1 To 1)
    MapInfo(1).Seguro = False
    MapInfo(1).SafeFightMap = False

    MapData(1, 10, 10).NpcIndex = npcIndex
    MapData(1, 9, 10).Blocked = e_Block.ALL_SIDES
    MapData(1, 11, 10).Blocked = e_Block.ALL_SIDES
    MapData(1, 10, 9).Blocked = e_Block.ALL_SIDES
    MapData(1, 10, 11).Blocked = e_Block.ALL_SIDES

    With UserList(1)
        .VersionId = 1
        .flags.Privilegios = e_PlayerType.user
        .flags.Muerto = 0
        .flags.Montado = 0
        .flags.Inmunidad = 0
        .flags.EnConsulta = False
        .flags.Seguro = False
        .flags.CurrentTeam = 0
        .flags.AdminInvisible = 0
        .Grupo.EnGrupo = False
        .Grupo.Id = 0
        .GuildIndex = 0
        .Faccion.Status = e_Facciones.Criminal
        .Pos.Map = 1
        .Pos.X = 12
        .Pos.Y = 10
    End With
    LastUser = 1
    MapData(1, 12, 10).UserIndex = 1

    Debug.Assert SetUserRef(NpcList(npcIndex).TargetUser, 1)

    Call AI_CaminarConRumbo(npcIndex, UserList(1).Pos)

    Debug.Assert NpcList(npcIndex).pathFindingInfo.TargetUnreachable
    Debug.Assert NpcList(npcIndex).Attackable = 0
    attackCheck = UserCanAttackNpc(1, npcIndex)
    Debug.Assert attackCheck.Result = eInmuneNpc

    MapData(1, 12, 10).UserIndex = 0
    UserList(1).Pos.X = 10
    UserList(1).Pos.Y = 10
    MapData(1, 10, 10).UserIndex = 1
    MapData(1, 9, 10).Blocked = 0

    Call AI_CaminarConRumbo(npcIndex, UserList(1).Pos)

    Debug.Assert Not NpcList(npcIndex).pathFindingInfo.TargetUnreachable
    Debug.Assert NpcList(npcIndex).Attackable = 1
    attackCheck = UserCanAttackNpc(1, npcIndex)
    Debug.Assert attackCheck.Result = eCanAttack

    Call ResetNpcMainInfo(npcIndex)
    Call ResetNpcFlags(npcIndex)
    Call ResetNpcCounters(npcIndex)
    ReDim NpcList(npcIndex).pathFindingInfo.Path(1 To MAX_PATH_LENGTH)

    With NpcList(npcIndex)
        .Attackable = 1
        .Hostile = 1
        .Pos.Map = 1
        .Pos.X = 10
        .Pos.Y = 10
        .Orig = .Pos
        .AttackRange = 2
        .pathFindingInfo.RangoVision = 1
        .pathFindingInfo.OriginalVision = 1
        .pathFindingInfo.PathLength = 0
        .flags.Faccion = e_Facciones.Ciudadano
        .Humanoide = False
    End With

    NumMaps = 1
    MinXBorder = XMinMapSize
    MaxXBorder = XMaxMapSize
    MinYBorder = YMinMapSize
    MaxYBorder = YMaxMapSize

    ReDim MapData(1 To 1, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize)
    ReDim MapInfo(1 To 1)
    MapInfo(1).Seguro = False
    MapInfo(1).SafeFightMap = False

    MapData(1, 10, 10).NpcIndex = npcIndex
    MapData(1, 11, 10).Blocked = e_Block.ALL_SIDES
    MapData(1, 12, 10).Blocked = e_Block.ALL_SIDES

    With UserList(1)
        .VersionId = 1
        .flags.Privilegios = e_PlayerType.user
        .flags.Muerto = 0
        .flags.Montado = 0
        .flags.Inmunidad = 0
        .flags.EnConsulta = False
        .flags.Seguro = False
        .flags.CurrentTeam = 0
        .flags.AdminInvisible = 0
        .Grupo.EnGrupo = False
        .Grupo.Id = 0
        .GuildIndex = 0
        .Faccion.Status = e_Facciones.Criminal
        .Pos.Map = 1
        .Pos.X = 12
        .Pos.Y = 10
    End With
    MapData(1, 12, 10).UserIndex = 1
    LastUser = 1

    Debug.Assert SetUserRef(NpcList(npcIndex).TargetUser, 1)

    Call AI_CaminarConRumbo(npcIndex, UserList(1).Pos)

    Debug.Assert Not NpcList(npcIndex).pathFindingInfo.TargetUnreachable
    Debug.Assert NpcList(npcIndex).Attackable = 1

    test_npc_pathfinding_attackable_state = True
End Function

Function test_suite() As Boolean
    Dim result As Boolean
    result = test_make_user_char()
    result = result And test_maths()
    result = result And test_npc_pathfinding_attackable_state()
    test_suite = result
End Function

#End If
