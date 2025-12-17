Attribute VB_Name = "General"
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
Public Declare Function QueryPerformanceCounter Lib "Kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "Kernel32" (lpFrequency As Currency) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Sub OutputDebugString Lib "Kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Global LeerNPCs As New clsIniManager

Private Const TREE_GRAPHICS_FILE As String = "EsArbol.ini"
Private TreeGraphicIds()         As Long
Private TreeGraphicCount         As Long

Sub SetNakedBody(ByRef User As t_User)
    Const man_human_naked_body   As Integer = 3000
    Const man_drow_naked_body    As Integer = 3001
    Const man_elf_naked_body     As Integer = 3002
    Const man_gnome_naked_body   As Integer = 3003
    Const man_dwarf_naked_body   As Integer = 3004
    Const man_orc_naked_body     As Integer = 3005
    Const woman_human_naked_body As Integer = 3006
    Const woman_drow_naked_body  As Integer = 3007
    Const woman_elf_naked_body   As Integer = 3008
    Const woman_gnome_naked_body As Integer = 3009
    Const woman_dwarf_naked_body As Integer = 3010
    Const woman_orc_naked_body   As Integer = 3011
    User.flags.Desnudo = 1
    Select Case User.genero
        Case e_Genero.Hombre
            Select Case User.raza
                Case e_Raza.Humano
                    User.Char.body = man_human_naked_body
                Case e_Raza.Drow
                    User.Char.body = man_drow_naked_body
                Case e_Raza.Elfo
                    User.Char.body = man_elf_naked_body
                Case e_Raza.Gnomo
                    User.Char.body = man_gnome_naked_body
                Case e_Raza.Enano
                    User.Char.body = man_dwarf_naked_body
                Case e_Raza.Orco
                    User.Char.body = man_orc_naked_body
                Case Else
                    User.Char.body = man_human_naked_body
            End Select
        Case e_Genero.Mujer
            Select Case User.raza
                Case e_Raza.Humano
                    User.Char.body = woman_human_naked_body
                Case e_Raza.Drow
                    User.Char.body = woman_drow_naked_body
                Case e_Raza.Elfo
                    User.Char.body = woman_elf_naked_body
                Case e_Raza.Gnomo
                    User.Char.body = woman_gnome_naked_body
                Case e_Raza.Enano
                    User.Char.body = woman_dwarf_naked_body
                Case e_Raza.Orco
                    User.Char.body = woman_orc_naked_body
                Case Else
                    User.Char.body = woman_human_naked_body
            End Select
    End Select
End Sub

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal b As Byte)
    'b ahora es boolean,
    'b=true bloquea el tile en (x,y)
    'b=false desbloquea el tile en (x,y)
    'toMap = true -> Envia los datos a todo el mapa
    'toMap = false -> Envia los datos al user
    'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
    'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
    '  Uso bloqueo parcial
    On Error GoTo Bloquear_Err
    ' Envío sólo los flags de bloq
    b = b And e_Block.ALL_SIDES
    If toMap Then
        Call SendData(SendTarget.toMap, sndIndex, PrepareMessage_BlockPosition(x, y, b))
    Else
        Call Write_BlockPosition(sndIndex, x, y, b)
    End If
    Exit Sub
Bloquear_Err:
    Call TraceError(Err.Number, Err.Description, "General.Bloquear", Erl)
End Sub

Sub BlockAndInform(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal NewState As Integer)
    If NewState Then
        MapData(Map, x, y).Blocked = e_Block.ALL_SIDES Or e_Block.GM
    Else
        MapData(Map, x, y).Blocked = 0
    End If
    Call Bloquear(True, Map, x, y, MapData(Map, x, y).Blocked)
End Sub

Sub MostrarBloqueosPuerta(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo MostrarBloqueosPuerta_Err
    Dim Map       As Integer
    Dim ModPuerta As Integer
    If toMap Then
        Map = sndIndex
    Else
        Map = UserList(sndIndex).pos.Map
    End If
    ModPuerta = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Subtipo
    Select Case ModPuerta
        Case 0
            ' Bloqueos superiores
            Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y, MapData(Map, x - 1, y).Blocked)
            ' Bloqueos inferiores
            Call Bloquear(toMap, sndIndex, x, y + 1, MapData(Map, x, y + 1).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y + 1, MapData(Map, x - 1, y + 1).Blocked)
        Case 1
            ' para palancas o teclas sin modicar bloqueos en X,Y
        Case 2
            ' Bloqueos superiores
            Call Bloquear(toMap, sndIndex, x, y - 1, MapData(Map, x, y - 1).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y - 1, MapData(Map, x - 1, y - 1).Blocked)
            Call Bloquear(toMap, sndIndex, x + 1, y - 1, MapData(Map, x + 1, y - 1).Blocked)
            ' Bloqueos inferiores
            Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y, MapData(Map, x - 1, y).Blocked)
            Call Bloquear(toMap, sndIndex, x + 1, y, MapData(Map, x + 1, y).Blocked)
        Case 3
            ' Bloqueos superiores
            Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y, MapData(Map, x - 1, y).Blocked)
            Call Bloquear(toMap, sndIndex, x + 1, y, MapData(Map, x + 1, y).Blocked)
            ' Bloqueos inferiores
            Call Bloquear(toMap, sndIndex, x, y + 1, MapData(Map, x, y + 1).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y + 1, MapData(Map, x - 1, y + 1).Blocked)
            Call Bloquear(toMap, sndIndex, x + 1, y + 1, MapData(Map, x + 1, y + 1).Blocked)
        Case 4
            ' Bloqueos superiores
            Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
            ' Bloqueos inferiores
            Call Bloquear(toMap, sndIndex, x, y + 1, MapData(Map, x, y + 1).Blocked)
        Case 5 'Ver WyroX
            ' Bloqueos vertical ver ReyarB
            Call Bloquear(toMap, sndIndex, x + 1, y, MapData(Map, x + 1, y).Blocked)
            Call Bloquear(toMap, sndIndex, x + 1, y - 1, MapData(Map, x + 1, y - 1).Blocked)
            ' Bloqueos horizontal
            Call Bloquear(toMap, sndIndex, x, y - 2, MapData(Map, x, y - 2).Blocked)
            Call Bloquear(toMap, sndIndex, x - 1, y - 2, MapData(Map, x - 1, y - 2).Blocked)
        Case 6 ' Ver WyroX
            ' Bloqueos superiores ver ReyarB
            Call Bloquear(toMap, sndIndex, x, y, MapData(Map, x, y).Blocked)
            Call Bloquear(toMap, sndIndex, x, y - 1, MapData(Map, x, y - 1).Blocked)
            ' Bloqueos inferiores
            Call Bloquear(toMap, sndIndex, x, y - 2, MapData(Map, x, y - 2).Blocked)
            Call Bloquear(toMap, sndIndex, x + 1, y - 2, MapData(Map, x + 1, y - 2).Blocked)
    End Select
    Exit Sub
MostrarBloqueosPuerta_Err:
    Call TraceError(Err.Number, Err.Description, "General.MostrarBloqueosPuerta", Erl)
End Sub

Sub BloquearPuerta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Bloquear As Boolean)
    On Error GoTo BloquearPuerta_Err
    Dim ModPuerta As Integer
    'ver reyarb
    ModPuerta = ObjData(MapData(Map, x, y).ObjInfo.ObjIndex).Subtipo
    Select Case ModPuerta
        Case 0 'puerta 2 tiles
            ' Bloqueos superiores
            MapData(Map, x, y).Blocked = IIf(Bloquear, MapData(Map, x, y).Blocked Or e_Block.NORTH, MapData(Map, x, y).Blocked And Not e_Block.NORTH)
            MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.NORTH, MapData(Map, x - 1, y).Blocked And Not e_Block.NORTH)
            ' Cambio bloqueos inferiores
            MapData(Map, x, y + 1).Blocked = IIf(Bloquear, MapData(Map, x, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x, y + 1).Blocked And Not e_Block.SOUTH)
            MapData(Map, x - 1, y + 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x - 1, y + 1).Blocked And Not e_Block.SOUTH)
        Case 1
            ' para palancas o teclas sin modicar bloqueos en X,Y
        Case 2 ' puerta 3 tiles 1 arriba
            ' Bloqueos superiores
            MapData(Map, x, y - 1).Blocked = IIf(Bloquear, MapData(Map, x, y - 1).Blocked Or e_Block.NORTH, MapData(Map, x, y - 1).Blocked And Not e_Block.NORTH)
            MapData(Map, x - 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y - 1).Blocked Or e_Block.NORTH, MapData(Map, x - 1, y - 1).Blocked And Not e_Block.NORTH)
            MapData(Map, x + 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x + 1, y - 1).Blocked Or e_Block.NORTH, MapData(Map, x + 1, y - 1).Blocked And Not e_Block.NORTH)
            ' Cambio bloqueos inferiores
            MapData(Map, x, y).Blocked = IIf(Bloquear, MapData(Map, x, y).Blocked Or e_Block.SOUTH, MapData(Map, x, y).Blocked And Not e_Block.SOUTH)
            MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.SOUTH, MapData(Map, x - 1, y).Blocked And Not e_Block.SOUTH)
            MapData(Map, x + 1, y).Blocked = IIf(Bloquear, MapData(Map, x + 1, y).Blocked Or e_Block.SOUTH, MapData(Map, x + 1, y).Blocked And Not e_Block.SOUTH)
        Case 3 ' puerta 3 tiles
            ' Bloqueos superiores
            MapData(Map, x, y).Blocked = IIf(Bloquear, MapData(Map, x, y).Blocked Or e_Block.NORTH, MapData(Map, x, y).Blocked And Not e_Block.NORTH)
            MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.NORTH, MapData(Map, x - 1, y).Blocked And Not e_Block.NORTH)
            MapData(Map, x + 1, y).Blocked = IIf(Bloquear, MapData(Map, x + 1, y).Blocked Or e_Block.NORTH, MapData(Map, x + 1, y).Blocked And Not e_Block.NORTH)
            ' Cambio bloqueos inferiores
            MapData(Map, x, y + 1).Blocked = IIf(Bloquear, MapData(Map, x, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x, y + 1).Blocked And Not e_Block.SOUTH)
            MapData(Map, x - 1, y + 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x - 1, y + 1).Blocked And Not e_Block.SOUTH)
            MapData(Map, x + 1, y + 1).Blocked = IIf(Bloquear, MapData(Map, x + 1, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x + 1, y + 1).Blocked And Not e_Block.SOUTH)
        Case 4 'puerta 1 tiles
            ' Bloqueos superiores
            MapData(Map, x, y).Blocked = IIf(Bloquear, MapData(Map, x, y).Blocked Or e_Block.NORTH, MapData(Map, x, y).Blocked And Not e_Block.NORTH)
            ' Cambio bloqueos inferiores
            MapData(Map, x, y + 1).Blocked = IIf(Bloquear, MapData(Map, x, y + 1).Blocked Or e_Block.SOUTH, MapData(Map, x, y + 1).Blocked And Not e_Block.SOUTH)
        Case 5 'Ver WyroX
            ' Bloqueos  vertical ver ReyarB
            MapData(Map, x + 1, y).Blocked = IIf(Bloquear, MapData(Map, x + 1, y).Blocked Or e_Block.ALL_SIDES, MapData(Map, x + 1, y).Blocked And Not e_Block.ALL_SIDES)
            MapData(Map, x + 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x + 1, y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, x + 1, y - 1).Blocked And Not _
                    e_Block.ALL_SIDES)
            ' Cambio horizontal
            MapData(Map, x, y - 2).Blocked = IIf(Bloquear, MapData(Map, x, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x, y - 2).Blocked And Not e_Block.ALL_SIDES)
            MapData(Map, x - 1, y - 2).Blocked = IIf(Bloquear, MapData(Map, x - 1, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x - 1, y - 2).Blocked And Not _
                    e_Block.ALL_SIDES)
        Case 6 ' Ver Wyrox
            ' Bloqueos vertical ver ReyarB
            MapData(Map, x - 1, y).Blocked = IIf(Bloquear, MapData(Map, x - 1, y).Blocked Or e_Block.ALL_SIDES, MapData(Map, x - 1, y).Blocked And Not e_Block.ALL_SIDES)
            MapData(Map, x - 1, y - 1).Blocked = IIf(Bloquear, MapData(Map, x - 1, y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, x - 1, y - 1).Blocked And Not _
                    e_Block.ALL_SIDES)
            ' Cambio bloqueos Puerta abierta
            MapData(Map, x, y - 2).Blocked = IIf(Bloquear, MapData(Map, x, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x, y - 2).Blocked And Not e_Block.ALL_SIDES)
            MapData(Map, x + 1, y + 2).Blocked = IIf(Bloquear, MapData(Map, x + 1, y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, x + 1, y - 2).Blocked And Not _
                    e_Block.ALL_SIDES)
    End Select
    ' Mostramos a todos
    Call MostrarBloqueosPuerta(True, Map, x, y)
    Exit Sub
BloquearPuerta_Err:
    Call TraceError(Err.Number, Err.Description, "General.BloquearPuerta", Erl)
End Sub

Function HayCosta(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo HayCosta_Err
    'Ladder 10 - 2 - 2010
    'Chequea si hay costa en los tiles proximos al usuario
    If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
        If ((MapData(Map, x, y).Graphic(1) >= 22552 And MapData(Map, x, y).Graphic(1) <= 22599) Or (MapData(Map, x, y).Graphic(1) >= 7283 And MapData(Map, x, y).Graphic(1) <= _
                7378) Or (MapData(Map, x, y).Graphic(1) >= 13387 And MapData(Map, x, y).Graphic(1) <= 13482)) And MapData(Map, x, y).Graphic(2) = 0 Then
            HayCosta = True
        Else
            HayCosta = False
        End If
    Else
        HayCosta = False
    End If
    Exit Function
HayCosta_Err:
    Call TraceError(Err.Number, Err.Description, "General.HayCosta", Erl)
End Function

Function HayAgua(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo HayAgua_Err
    With MapData(Map, x, y)
        If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
            HayAgua = (.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or (.Graphic(1) >= 124 And .Graphic(1) <= 139) Or (.Graphic(1) >= 24223 And .Graphic(1) <= 24238) Or ( _
                    .Graphic(1) >= 24303 And .Graphic(1) <= 24318) Or (.Graphic(1) >= 468 And .Graphic(1) <= 483) Or (.Graphic(1) >= 44668 And .Graphic(1) <= 44683) Or (.Graphic( _
                    1) >= 24143 And .Graphic(1) <= 24158) Or (.Graphic(1) >= 12628 And .Graphic(1) <= 12643) Or (.Graphic(1) >= 2948 And .Graphic(1) <= 2963)
        Else
            HayAgua = False
        End If
    End With
    Exit Function
HayAgua_Err:
    Call TraceError(Err.Number, Err.Description, "General.HayAgua", Erl)
End Function

Public Sub LoadTreeGraphics()
    On Error GoTo LoadTreeGraphics_Err
    Dim initTreeFile   As String
    Dim legacyTreeFile As String

    initTreeFile = GetTreeGraphicsInitPath()
    legacyTreeFile = GetTreeGraphicsLegacyPath()

    If LoadTreeGraphicsFromFile(initTreeFile) Then Exit Sub

    If LoadTreeGraphicsFromFile(legacyTreeFile) Then Exit Sub

    Call LoadDefaultTreeGraphics

    If EnsureDirectoryExists(GetDirectoryName(initTreeFile)) Then
        Call SaveTreeGraphicsFile(initTreeFile)
    ElseIf EnsureDirectoryExists(GetDirectoryName(legacyTreeFile)) Then
        Call SaveTreeGraphicsFile(legacyTreeFile)
    Else
        Call SaveTreeGraphicsFile(initTreeFile)
    End If
    Exit Sub
LoadTreeGraphics_Err:
    Call TraceError(Err.Number, Err.Description, "General.LoadTreeGraphics", Erl)
End Sub

Private Function GetTreeGraphicsInitPath() As String
    GetTreeGraphicsInitPath = App.Path & "\Recursos\init\" & TREE_GRAPHICS_FILE
End Function

Private Function GetTreeGraphicsLegacyPath() As String
    GetTreeGraphicsLegacyPath = DatPath & TREE_GRAPHICS_FILE
End Function


Private Function LoadTreeGraphicsFromFile(ByVal FilePath As String) As Boolean
    On Error GoTo LoadTreeGraphicsFromFile_Err

    If Not FileExist(FilePath, vbArchive) Then Exit Function

    Dim requestedCount As Long
    Dim idx            As Long
    Dim treeValue      As Long

    requestedCount = val(GetVar(FilePath, "INIT", "Count"))

    If requestedCount <= 0 Then Exit Function

    ReDim TreeGraphicIds(1 To requestedCount) As Long
    TreeGraphicCount = 0

    For idx = 1 To requestedCount
        treeValue = val(GetVar(FilePath, "TREES", "Tree" & idx))
        If treeValue <> 0 Then
            TreeGraphicCount = TreeGraphicCount + 1
            TreeGraphicIds(TreeGraphicCount) = treeValue
        End If
    Next idx

    If TreeGraphicCount = 0 Then
        Erase TreeGraphicIds
        Exit Function
    End If

    If TreeGraphicCount <> requestedCount Then
        ReDim Preserve TreeGraphicIds(1 To TreeGraphicCount) As Long
    End If

    LoadTreeGraphicsFromFile = True
    Exit Function
LoadTreeGraphicsFromFile_Err:
    Call TraceError(Err.Number, Err.Description, "General.LoadTreeGraphicsFromFile", Erl)
End Function

Private Sub LoadDefaultTreeGraphics()
    On Error GoTo LoadDefaultTreeGraphics_Err

    Dim defaults As Variant
    Dim idx      As Long
    Dim dest     As Long

    defaults = Array(11905&, 644&, 1880&, 11906&, 12160&, 6597&, 2548&, 2549&, 15110&, 15109&, 15108&, 11904&, 7220&, 50990&, 55626&, 55627&, 55630&, 55632&, 55633&, 55635&, 55638&, 12584&, 50985&, 15510&, 14775&, 14687&, 11903&, 735&, 15698&, 14504&, 15697&, 6598&, 1121&, 1878&, 9513&, 9514&, 9515&, 9518&, 9519&, 9520&, 9529&)

    TreeGraphicCount = UBound(defaults) - LBound(defaults) + 1
    ReDim TreeGraphicIds(1 To TreeGraphicCount) As Long

    dest = 1
    For idx = LBound(defaults) To UBound(defaults)
        TreeGraphicIds(dest) = CLng(defaults(idx))
        dest = dest + 1
    Next idx
    Exit Sub
LoadDefaultTreeGraphics_Err:
    Call TraceError(Err.Number, Err.Description, "General.LoadDefaultTreeGraphics", Erl)
End Sub

Private Sub SaveTreeGraphicsFile(ByVal FilePath As String)
    On Error GoTo SaveTreeGraphicsFile_Err

    Dim idx As Long

    If TreeGraphicCount = 0 Then Exit Sub

    Call WriteVar(FilePath, "INIT", "Count", CStr(TreeGraphicCount))

    For idx = 1 To TreeGraphicCount
        Call WriteVar(FilePath, "TREES", "Tree" & idx, CStr(TreeGraphicIds(idx)))
    Next idx
    Exit Sub
SaveTreeGraphicsFile_Err:
    Call TraceError(Err.Number, Err.Description, "General.SaveTreeGraphicsFile", Erl)
End Sub

Function EsArbol(ByVal GrhIndex As Long) As Boolean
    On Error GoTo EsArbol_Err

    Dim idx As Long

    For idx = 1 To TreeGraphicCount
        If TreeGraphicIds(idx) = GrhIndex Then
            EsArbol = True
            Exit Function
        End If
    Next idx
    Exit Function
EsArbol_Err:
    Call TraceError(Err.Number, Err.Description, "General.EsArbol", Erl)
End Function

Private Function EnsureDirectoryExists(ByVal DirectoryPath As String) As Boolean
    On Error GoTo EnsureDirectoryExists_Err

    Dim normalizedPath As String
    Dim parentDirectory As String

    normalizedPath = NormalizePath(DirectoryPath)

    If LenB(normalizedPath) = 0 Then Exit Function

    If Len(normalizedPath) = 2 And Mid$(normalizedPath, 2, 1) = ":" Then
        EnsureDirectoryExists = True
        Exit Function
    End If

    If FileExist(normalizedPath, vbDirectory) Then
        EnsureDirectoryExists = True
        Exit Function
    End If

    parentDirectory = GetDirectoryName(normalizedPath)
    If LenB(parentDirectory) <> 0 Then
        If Not EnsureDirectoryExists(parentDirectory) Then Exit Function
    End If

    MkDir normalizedPath
    EnsureDirectoryExists = True
    Exit Function
EnsureDirectoryExists_Err:
    EnsureDirectoryExists = False
End Function

Private Function GetDirectoryName(ByVal PathValue As String) As String
    Dim normalizedPath As String
    Dim separatorPos   As Long

    normalizedPath = NormalizePath(PathValue)
    separatorPos = InStrRev(normalizedPath, "\")

    If separatorPos > 0 Then
        GetDirectoryName = Left$(normalizedPath, separatorPos - 1)
    End If
End Function

Private Function NormalizePath(ByVal PathValue As String) As String
    Dim normalizedPath As String

    normalizedPath = Replace$(PathValue, "/", "\")

    Do While Len(normalizedPath) > 0 And Right$(normalizedPath, 1) = "\"
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    Loop

    NormalizePath = normalizedPath
End Function

Private Function HayLava(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo HayLava_Err
    If Map > 0 And Map < NumMaps + 1 And x > 0 And x < 101 And y > 0 And y < 101 Then
        If MapData(Map, x, y).Graphic(1) >= 5837 And MapData(Map, x, y).Graphic(1) <= 5852 Or MapData(Map, x, y).Graphic(1) >= 16101 And MapData(Map, x, y).Graphic(1) <= 16116 Then
            HayLava = True
        Else
            HayLava = False
        End If
    Else
        HayLava = False
    End If
    Exit Function
HayLava_Err:
    Call TraceError(Err.Number, Err.Description, "General.HayLava", Erl)
End Function

Sub ApagarFogatas()
    'Ladder /ApagarFogatas
    On Error GoTo ErrHandler
    Dim obj As t_Obj
    obj.ObjIndex = FOGATA_APAG
    obj.amount = 1
    Dim MapaActual As Long
    Dim y          As Long
    Dim x          As Long
    For MapaActual = 1 To NumMaps
        For y = YMinMapSize To YMaxMapSize
            For x = XMinMapSize To XMaxMapSize
                If MapInfo(MapaActual).lluvia Then
                    If MapData(MapaActual, x, y).ObjInfo.ObjIndex = FOGATA Then
                        Call EraseObj(GetMaxInvOBJ(), MapaActual, x, y)
                        Call MakeObj(obj, MapaActual, x, y)
                    End If
                End If
            Next x
        Next y
    Next MapaActual
    Exit Sub
ErrHandler:
    Call LogError("Error producido al apagar las fogatas de " & x & "-" & y & " del mapa: " & MapaActual & "    -" & Err.Description)
End Sub

Private Sub InicializarConstantes()
    On Error GoTo InicializarConstantes_Err
    LastBackup = Format$(Now, "Short Time")
    minutos = Format$(Now, "Short Time")
    IniPath = App.Path & "\"
    ListaRazas(e_Raza.Humano) = "Humano"
    ListaRazas(e_Raza.Elfo) = "Elfo"
    ListaRazas(e_Raza.Drow) = "Elfo Oscuro"
    ListaRazas(e_Raza.Gnomo) = "Gnomo"
    ListaRazas(e_Raza.Enano) = "Enano"
    ListaRazas(e_Raza.Orco) = "Orco"
    ListaClases(e_Class.Mage) = "Mago"
    ListaClases(e_Class.Cleric) = "Clérigo"
    ListaClases(e_Class.Warrior) = "Guerrero"
    ListaClases(e_Class.Assasin) = "Asesino"
    ListaClases(e_Class.Bard) = "Bardo"
    ListaClases(e_Class.Druid) = "Druida"
    ListaClases(e_Class.Paladin) = "Paladín"
    ListaClases(e_Class.Hunter) = "Cazador"
    ListaClases(e_Class.Trabajador) = "Trabajador"
    ListaClases(e_Class.Pirat) = "Pirata"
    ListaClases(e_Class.Thief) = "Ladrón"
    ListaClases(e_Class.Bandit) = "Bandido"
    SkillsNames(e_Skill.Magia) = "Magia"
    SkillsNames(e_Skill.Robar) = "Robar"
    SkillsNames(e_Skill.Tacticas) = "Destreza en combate"
    SkillsNames(e_Skill.Armas) = "Combate con armas"
    SkillsNames(e_Skill.Meditar) = "Meditar"
    SkillsNames(e_Skill.Apuñalar) = "Apuñalar"
    SkillsNames(e_Skill.Ocultarse) = "Ocultarse"
    SkillsNames(e_Skill.Supervivencia) = "Supervivencia"
    SkillsNames(e_Skill.Comerciar) = "Comercio"
    SkillsNames(e_Skill.Defensa) = "Defensa con escudo"
    SkillsNames(e_Skill.liderazgo) = "Liderazgo"
    SkillsNames(e_Skill.Proyectiles) = "Armas a distancia"
    SkillsNames(e_Skill.Wrestling) = "Combate sin armas"
    SkillsNames(e_Skill.Navegacion) = "Navegación"
    SkillsNames(e_Skill.equitacion) = "Equitación"
    SkillsNames(e_Skill.Resistencia) = "Resistencia mágica"
    SkillsNames(e_Skill.Talar) = "Tala"
    SkillsNames(e_Skill.Pescar) = "Pesca"
    SkillsNames(e_Skill.Mineria) = "Minería"
    SkillsNames(e_Skill.Herreria) = "Herrería"
    SkillsNames(e_Skill.Carpinteria) = "Carpintería"
    SkillsNames(e_Skill.Alquimia) = "Alquimia"
    SkillsNames(e_Skill.Sastreria) = "Sastrería"
    SkillsNames(e_Skill.Domar) = "Domar"
    ListaAtributos(e_Atributos.Fuerza) = "Fuerza"
    ListaAtributos(e_Atributos.Agilidad) = "Agilidad"
    ListaAtributos(e_Atributos.Inteligencia) = "Inteligencia"
    ListaAtributos(e_Atributos.Constitucion) = "Constitución"
    ListaAtributos(e_Atributos.Carisma) = "Carisma"
    IniPath = App.Path & "\"
    'Bordes del mapa
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    RaceHeightOffset(Humano) = -35
    RaceHeightOffset(Elfo) = -35
    RaceHeightOffset(Drow) = -35
    RaceHeightOffset(Gnomo) = -27
    RaceHeightOffset(Enano) = -27
    RaceHeightOffset(Orco) = -35
    WeaponTypeNames(eSword) = "Sword"
    WeaponTypeNames(eDagger) = "Dagger"
    WeaponTypeNames(eBow) = "Bow"
    WeaponTypeNames(eStaff) = "Staff"
    WeaponTypeNames(eMace) = "Mace"
    WeaponTypeNames(eThrowableAxe) = "ThrowableAxe"
    WeaponTypeNames(eAxe) = "Axe"
    WeaponTypeNames(eKnuckle) = "Knuckle"
    WeaponTypeNames(e_WeaponType.eFist) = "Fist"
    WeaponTypeNames(e_WeaponType.eSpear) = "Spear"
    WeaponTypeNames(e_WeaponType.eGunPowder) = "GunPowder"
    Exit Sub
InicializarConstantes_Err:
    Call TraceError(Err.Number, Err.Description, "General.InicializarConstantes", Erl)
End Sub

Sub Main()
    On Error GoTo Handler
        
    Call TryInitShard
    
    Call Uptime_Init
    #If DIRECT_PLAY = 1 Then
        InitDPlay
    #End If
    Call InitializeCircularLogBuffer
    Call LogThis(0, "Starting the server " & Now, vbLogEventTypeInformation)
    Call load_stats
    
    If Not IsShardingEnabled() Then
        If GetProcessCount(App.EXEName & ".exe") > 1 Then
            ' Si lo hay, pregunto si lo queremos cerrar.
            If MsgBox("Se ha encontrado mas de 1 instancia abierta de esta aplicación, ¿Desea continuar?", vbYesNo) = vbNo Then
                End
            End If
        End If
        frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    Else
        frmMain.Caption = ShardID & " -> " & frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    End If
    
    Call ChDir(App.Path)
    Call ChDrive(App.Path)
    Call InicializarConstantes
    frmCargando.Show
    
    frmCargando.Label1(2).Caption = "Iniciando Arrays..."
    Call InitializeNpcIndexHeap
    Call InitializeLobbyList
    Call loadAdministrativeUsers
    '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
    frmCargando.Label1(2).Caption = "Cargando Server.ini"
    MaxUsers = 0
    Call LoadSini
    Call LoadMD5
    Call LoadPacketRatePolicy
    #If PYMMO = 1 Then
        Call LoadPrivateKey
    #End If
    Call LoadMainConfigFile
    Call LoadIntervalos
    Call CargarForbidenWords
    Call LoadBlockedWordsDescription
    Call CargaApuestas
    Call CargarSpawnList
    Call LoadMotd
    Call initBase64Chars
    frmCargando.Label1(2).Caption = "Conectando base de datos y limpiando usuarios logueados"
    If Not FileExist(App.Path & "/" & DatabaseFileName) Then
        Call FileSystem.FileCopy(App.Path & "/Empty_db.db", App.Path & "/" & DatabaseFileName)
    End If
    ' ************************* Base de Datos ********************
    'Conecto base de datos
    Call Database_Connect
    Call Database_Connect_Async
    ' Construimos las querys grandes
    Call Contruir_Querys
    Call LoadDBMigrations
    ' ******************* FIN - Base de Datos ********************
    Call LoadGuildsDB
    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
    Call CargaNpcsDat
    '*************************************************
    frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
    Call LoadOBJData
    frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
    Call CargarHechizos
    frmCargando.Label1(2).Caption = "Cargando EffectsOverTime.Dat"
    Call LoadEffectOverTime
    frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
    Call LoadArmasHerreria
    Call LoadArmadurasHerreria
    Call LoadBlackSmithElementalRunes
    frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
    Call LoadObjCarpintero
    frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
    Call LoadObjAlquimista
    frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"
    Call LoadObjSastre
    frmCargando.Label1(2).Caption = "Cargando Pesca"
    Call LoadPesca
    Call InitializeFishingBonuses()
    frmCargando.Label1(2).Caption = "Cargando Recursos Especiales"
    Call LoadRecursosEspeciales
    frmCargando.Label1(2).Caption = "Cargando definiciones de árboles"
    Call LoadTreeGraphics
    frmCargando.Label1(2).Caption = "Cargando Rangos de Faccion"
    Call LoadRangosFaccion
    frmCargando.Label1(2).Caption = "Cargando Recompensas de Faccion"
    Call LoadRecompensasFaccion
    frmCargando.Label1(2).Caption = "Cargando Balance.dat"
    Call LoadBalance
    frmCargando.Label1(2).Caption = "Cargando Clanes.dat"
    Call LoadGuildsConfig
    frmCargando.Label1(2).Caption = "Cargando Ciudades.dat"
    Call CargarCiudades
    If BootDelBackUp Then
        frmCargando.Label1(2).Caption = "Cargando WorldBackup"
        Call CargarBackUp
    Else
        frmCargando.Label1(2).Caption = "Cargando Mapas"
        Call LoadMapData
    End If
    frmCargando.Label1(2).Caption = "Cargando donadores"
    Call CargarDonadores
    Call InitPathFinding
    frmCargando.Label1(2).Caption = "Cargando informacion de eventos"
    Call CargarInfoRetos
    Call CargarInfoEventos
    frmCargando.Label1(2).Caption = "Cargando Baneos Temporales"
    Call LoadBans
    frmCargando.Label1(2).Caption = "Cargando Quests"
    Call LoadQuests
    Call ResetLastLogoutAndIsLogged
    'Comentado porque hay worldsave en ese mapa!
    Dim LoopC As Integer
    'Resetea las conexiones de los usuarios
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnectionDetails.ConnIDValida = False
    Next LoopC
    With frmMain
        .Minuto.Enabled = True
        .tPiqueteC.Enabled = True
        .Segundo.Enabled = True
        .KillLog.Enabled = True
        .T_UsersOnline.Enabled = True
        .t_Extraer.Enabled = True
        .t_Extraer.Interval = IntervaloTrabajarExtraer
        .tControlHechizos.Enabled = True
        .tControlHechizos.Interval = 60000
        If IsFeatureEnabled("ShipTravelEnabled") Then
            .TimerBarco.Enabled = True
            MapInfo(BarcoNavegandoForgatNix.Map).ForceUpdate = True
            MapInfo(BarcoNavegandoNixArghal.Map).ForceUpdate = True
            MapInfo(BarcoNavegandoArghalForgat.Map).ForceUpdate = True
        End If
    End With
    Call ResetGameEventsTimer
    Call ResetUserAutoSaveTimer
    Subasta.SubastaHabilitada = True
    Subasta.HaySubastaActiva = False
    Call ResetMeteo
    #If DIRECT_PLAY = 0 Then
        Call Protocol_Writes.InitializeAuxiliaryBuffer
    #End If
    Call modNetwork.Listen(MaxUsers, ListenIp, CStr(Puerto))
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    ' ----------------------------------------------------
    '           Configuracion de los sockets
    ' ----------------------------------------------------
    Call GetHoraActual
    WorldTime_Init CLng(SvrConfig.GetValue("DayLength")), 0
    frmCargando.Visible = False
    Unload frmCargando
    'Ocultar
    Call frmMain.InitMain(HideMe)
    Call InitializeAntiCheat
    tInicioServer = GetTickCountRaw()
    #If UNIT_TEST = 1 Then
        Call UnitTesting.Init
        Debug.Print "AO20 Unit Testing"
        Dim suite_passed_ok As Boolean
        suite_passed_ok = UnitTesting.test_suite()
        If (suite_passed_ok) Then
            Debug.Print "suite_passed_ok!!!"
        Else
            Debug.Print "suite failed!!!"
        End If
        Debug.Assert (suite_passed_ok)
        Debug.Print "Running proto suite, trying to connect to 127.0.0.1:7667"
        Call UnitClient.Init
        Call UnitClient.Connect("127.0.0.1", "7667")
    #End If
    While (True)
        GlobalFrameTime = GetTickCountRaw()
        Dim PerformanceTimer As Long
        Call PerformanceTestStart(PerformanceTimer)
        #If PYMMO = 1 Then
            Call modNetwork.close_not_logged_sockets_if_timeout
        #End If
        Call PerformTimeLimitCheck(PerformanceTimer, "General modNetwork.close_not_logged_sockets_if_timeout")
        #If DIRECT_PLAY = 0 Then
            Call modNetwork.Tick(GetElapsed())
        #End If
        Call PerformTimeLimitCheck(PerformanceTimer, "General modNetwork.Tick")
        Call UpdateEffectOverTime
        Call PerformTimeLimitCheck(PerformanceTimer, "General Update Effects over time")
        Call MaybeRunGameEvents
        Call PerformTimeLimitCheck(PerformanceTimer, "General MaybeRunGameEvents")
        Call MaybeRunUserAutoSave
        Call PerformTimeLimitCheck(PerformanceTimer, "General MaybeRunUserAutoSave")
        Call RunAutomatedActions
        Call PerformTimeLimitCheck(PerformanceTimer, "General StartAutomatedAction")
        Call MaybeUpdateNpcAI(GlobalFrameTime)
        DoEvents
        Call PerformTimeLimitCheck(PerformanceTimer, "Do events")
        Call AntiCheatUpdate
        Call PerformTimeLimitCheck(PerformanceTimer, "Update anti cheat")
        ' Unlock main loop for maximum throughput but it can hog weak CPUs.
        #If UNLOCK_CPU = 0 Then
            Call Sleep(1)
        #End If
        #If UNIT_TEST = 1 Then
            Call UnitClient.Poll
        #End If
    Wend
    Call LogThis(0, "Closing the server " & Now, vbLogEventTypeInformation)
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "General.Main", Erl)
End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
    'Se fija si existe el archivo
    On Error GoTo FileExist_Err
    FileExist = LenB(dir$(File, FileType)) <> 0
    Exit Function
FileExist_Err:
    Call TraceError(Err.Number, Err.Description, "General.FileExist", Erl)
End Function

Function ReadField(ByVal pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
    On Error GoTo ReadField_Err
    'Gets a field from a delimited string
    Dim i          As Long
    Dim LastPos    As Long
    Dim currentPos As Long
    Dim delimiter  As String * 1
    delimiter = Chr$(SepASCII)
    For i = 1 To pos
        LastPos = currentPos
        currentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    If currentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, currentPos - LastPos - 1)
    End If
    Exit Function
ReadField_Err:
    Call TraceError(Err.Number, Err.Description, "General.ReadField", Erl)
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
    On Error GoTo MapaValido_Err
    MapaValido = Map >= 1 And Map <= NumMaps
    Exit Function
MapaValido_Err:
    Call TraceError(Err.Number, Err.Description, "General.MapaValido", Erl)
End Function

Sub MostrarNumUsers()
    On Error GoTo MostrarNumUsers_Err
    Call SendData(SendTarget.ToAll, 0, PrepareMessageOnlineUser(NumUsers))
    frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
    Exit Sub
MostrarNumUsers_Err:
    Call TraceError(Err.Number, Err.Description, "General.MostrarNumUsers", Erl)
End Sub

Sub Restart()
    On Error GoTo Restart_Err
    'Se asegura de que los sockets estan cerrados e ignora cualquier err
    If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."
    Dim LoopC As Long
    Call modNetwork.Disconnect
    For LoopC = 1 To MaxUsers
        Call CloseSocket(LoopC)
    Next
    'Initialize statistics!!
    'Call Statistics.Initialize
    ReDim UserList(1 To MaxUsers) As t_User
    For LoopC = 1 To MaxUsers
        UserList(LoopC).ConnectionDetails.ConnIDValida = False
    Next LoopC
    Call InitializeUserIndexHeap(MaxUsers)
    LastUser = 0
    NumUsers = 0
    Call FreeNPCs
    Call FreeCharIndexes
    Call LoadSini
    Call LoadMD5
    Call LoadPrivateKey
    Call LoadIntervalos
    Call ResetUserAutoSaveTimer
    Call LoadOBJData
    Call LoadPesca
    Call InitializeFishingBonuses()
    Call LoadRecursosEspeciales
    Call LoadTreeGraphics
    Call LoadMapData
    Call CargarHechizos
    Call modNetwork.Listen(MaxUsers, ListenIp, CStr(Puerto))
    If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
    'Ocultar
    Call frmMain.InitMain(HideMe)
    Exit Sub
Restart_Err:
    Call TraceError(Err.Number, Err.Description, "General.Restart", Erl)
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
    On Error GoTo Intemperie_Err
    If MapInfo(UserList(UserIndex).pos.Map).zone <> "DUNGEON" Then
        If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).trigger <> 1 And MapData(UserList(UserIndex).pos.Map, UserList( _
                UserIndex).pos.x, UserList(UserIndex).pos.y).trigger <> 2 And MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).trigger _
                < 10 Then Intemperie = True
    Else
        Intemperie = False
    End If
    Exit Function
Intemperie_Err:
    Call TraceError(Err.Number, Err.Description, "General.Intemperie", Erl)
End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
    On Error GoTo TiempoInvocacion_Err
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0 Then
            If Not IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
                Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
            Else
                If NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia > 0 Then
                    NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = NpcList(UserList(UserIndex).MascotasIndex( _
                            i).ArrayIndex).Contadores.TiempoExistencia - 1
                    If NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex( _
                            i).ArrayIndex, 0)
                End If
            End If
        End If
    Next i
    Exit Sub
TiempoInvocacion_Err:
    Call TraceError(Err.Number, Err.Description, "General.TiempoInvocacion", Erl)
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
    On Error GoTo EfectoFrio_Err
    If Not Intemperie(UserIndex) Then Exit Sub
    With UserList(UserIndex)
        If .invent.EquippedArmorObjIndex > 0 Then
            '  Ropa invernal
            If ObjData(.invent.EquippedArmorObjIndex).Invernal Then Exit Sub
        End If
        If .Counters.Frio < IntervaloFrio Then
            .Counters.Frio = .Counters.Frio + 1
        Else
            If MapInfo(.pos.Map).terrain = Nieve Then
                ' Msg512=¡Estás muriendo de frío, abrígate o morirás!
                Call WriteLocaleMsg(UserIndex, 512, e_FontTypeNames.FONTTYPE_INFO)
                '  Sin ropa perdés vida más rápido que con una ropa no-invernal
                Dim MinDamage As Integer, MaxDamage As Integer
                If .flags.Desnudo = 0 Then
                    MinDamage = 17
                    MaxDamage = 23
                Else
                    MinDamage = 27
                    MaxDamage = 33
                End If
                '  Agrego aleatoriedad
                Dim Damage As Integer
                Damage = Porcentaje(.Stats.MaxHp, RandomNumber(MinDamage, MaxDamage))
                If UserMod.ModifyHealth(UserIndex, -Damage, 0) Then
                    ' Msg513=¡Has muerto de frío!
                    Call WriteLocaleMsg(UserIndex, 513, e_FontTypeNames.FONTTYPE_INFO)
                    Call UserMod.UserDie(UserIndex)
                End If
            End If
            .Counters.Frio = 0
        End If
    End With
    Exit Sub
EfectoFrio_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoFrio", Erl)
End Sub

Public Sub EfectoStamina(ByVal UserIndex As Integer)
    Dim HambreOSed       As Boolean
    Dim bEnviarStats_HP  As Boolean
    Dim bEnviarStats_STA As Boolean
    With UserList(UserIndex)
        HambreOSed = .Stats.MinHam = 0 Or .Stats.MinAGU = 0
        'if hunger or thirst = 0 and not in combat
        If Not HambreOSed And .Counters.EnCombate = 0 Then
            If .Stats.MinHp < .Stats.MaxHp Then
                Call Sanar(UserIndex, bEnviarStats_HP, IIf(.flags.Descansar, SanaIntervaloDescansar, SanaIntervaloSinDescansar))
            End If
        End If
        If .flags.Desnudo = 0 And Not HambreOSed Then
            If (Not Lloviendo Or Not Intemperie(UserIndex)) And Not .AutomatedAction.IsActive Then
                Call RecStamina(UserIndex, bEnviarStats_STA, IIf(.flags.Descansar, StaminaIntervaloDescansar, StaminaIntervaloSinDescansar))
            End If
        Else
            If Lloviendo And Intemperie(UserIndex) Then
                Call PierdeEnergia(UserIndex, bEnviarStats_STA, IntervaloPerderStamina * 0.5)
            Else
                Call PierdeEnergia(UserIndex, bEnviarStats_STA, IIf(.flags.Descansar, IntervaloPerderStamina * 2, IntervaloPerderStamina))
            End If
        End If
        If .flags.Descansar Then
            'termina de descansar automaticamente
            If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                Call WriteRestOK(UserIndex)
                ' Msg514=Has terminado de descansar.
                Call WriteLocaleMsg(UserIndex, 514, e_FontTypeNames.FONTTYPE_INFO)
                .flags.Descansar = False
            End If
        End If
        If bEnviarStats_STA Then
            Call WriteUpdateSta(UserIndex)
        End If
        If bEnviarStats_HP Then
            Call WriteUpdateHP(UserIndex)
        End If
    End With
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
    On Error GoTo EfectoLava_Err
    With UserList(UserIndex)
        If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
            .Counters.Lava = .Counters.Lava + 1
        Else
            If HayLava(.pos.Map, .pos.x, .pos.y) Then
                ' Msg515=¡Quítate de la lava, te estás quemando!
                Call WriteLocaleMsg(UserIndex, 515, e_FontTypeNames.FONTTYPE_INFO)
                If UserMod.ModifyHealth(UserIndex, -Porcentaje(.Stats.MaxHp, 5)) Then
                    ' Msg516=¡Has muerto quemado!
                    Call WriteLocaleMsg(UserIndex, 516, e_FontTypeNames.FONTTYPE_INFO)
                    Call CustomScenarios.UserDie(UserIndex)
                    Call UserMod.UserDie(UserIndex)
                End If
            End If
            .Counters.Lava = 0
        End If
    End With
    Exit Sub
EfectoLava_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoLava", Erl)
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'
Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
    On Error GoTo EfectoMimetismo_Err
    Dim Barco As t_ObjData
    With UserList(UserIndex)
        If .Counters.Mimetismo < IntervaloInvisible Then
            .Counters.Mimetismo = .Counters.Mimetismo + 1
        Else
            'restore old char
            ' Msg517=Recuperas tu apariencia normal.
            Call WriteLocaleMsg(UserIndex, 517, e_FontTypeNames.FONTTYPE_INFO)
            If .flags.Navegando Then
                Call EquiparBarco(UserIndex)
            Else
                .Char.body = .CharMimetizado.body
                .Char.head = .CharMimetizado.head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Char.CartAnim = .CharMimetizado.CartAnim
            End If
            .Counters.Mimetismo = 0
            .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            With .Char
                Call ChangeUserChar(UserIndex, .body, .head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .CartAnim, .BackpackAnim)
                Call RefreshCharStatus(UserIndex)
            End With
        End If
    End With
    Exit Sub
EfectoMimetismo_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoMimetismo", Erl)
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
    On Error GoTo EfectoInvisibilidad_Err
    With UserList(UserIndex)
        If .Counters.Invisibilidad > 0 Then
            .Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad - 1
            If .Counters.DisabledInvisibility > 0 Then
                .Counters.DisabledInvisibility = .Counters.DisabledInvisibility - 1
                If .Counters.DisabledInvisibility = 0 And .Counters.Invisibilidad > 0 Then
                    .flags.invisible = 1
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, .pos.x, .pos.y))
                End If
            End If
        Else
            .Counters.Invisibilidad = 0
            .flags.invisible = 0
            .Counters.DisabledInvisibility = 0
            If .flags.Oculto = 0 Then
                ' Msg307=Has vuelto a ser visible
                Call WriteLocaleMsg(UserIndex, 307, e_FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, .pos.x, .pos.y))
                Call WriteContadores(UserIndex)
            End If
        End If
    End With
    Exit Sub
EfectoInvisibilidad_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoInvisibilidad", Erl)
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
    On Error GoTo EfectoParalisisNpc_Err
    If NpcList(NpcIndex).Contadores.Paralisis > 0 Then
        NpcList(NpcIndex).Contadores.Paralisis = NpcList(NpcIndex).Contadores.Paralisis - 1
    Else
        NpcList(NpcIndex).flags.Paralizado = 0
    End If
    Exit Sub
EfectoParalisisNpc_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoParalisisNpc", Erl)
End Sub

Public Sub EfectoInmovilizadoNpc(ByVal NpcIndex As Integer)
    On Error GoTo EfectoInmovilizadoNpc_Err
    If NpcList(NpcIndex).Contadores.Inmovilizado > 0 Then
        NpcList(NpcIndex).Contadores.Inmovilizado = NpcList(NpcIndex).Contadores.Inmovilizado - 1
    Else
        NpcList(NpcIndex).flags.Inmovilizado = 0
    End If
    Exit Sub
EfectoInmovilizadoNpc_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoInmovilizadoNpc", Erl)
End Sub

Public Sub EfectoCeguera(ByVal UserIndex As Integer)
    On Error GoTo EfectoCeguera_Err
    If UserList(UserIndex).Counters.Ceguera > 0 Then
        UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
    Else
        If UserList(UserIndex).flags.Ceguera = 1 Then
            UserList(UserIndex).flags.Ceguera = 0
            Call WriteBlindNoMore(UserIndex)
        End If
    End If
    Exit Sub
EfectoCeguera_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoCeguera", Erl)
End Sub

Public Sub EfectoEstupidez(ByVal UserIndex As Integer)
    On Error GoTo EfectoEstupidez_Err
    If UserList(UserIndex).Counters.Estupidez > 0 Then
        UserList(UserIndex).Counters.Estupidez = UserList(UserIndex).Counters.Estupidez - 1
    Else
        If UserList(UserIndex).flags.Estupidez = 1 Then
            UserList(UserIndex).flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
    End If
    Exit Sub
EfectoEstupidez_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoEstupidez", Erl)
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
    On Error GoTo EfectoParalisisUser_Err
    With UserList(UserIndex)
        If .Counters.Paralisis > 0 Then
            .Counters.Paralisis = .Counters.Paralisis - 1
        Else
            .flags.Paralizado = 0
            If .clase = e_Class.Warrior Or .clase = e_Class.Thief Or .clase = e_Class.Pirat Then
                .Counters.TiempoDeInmunidadParalisisNoMagicas = 3
            End If
            'UserList(UserIndex).Flags.AdministrativeParalisis = 0
            Call WriteParalizeOK(UserIndex)
        End If
    End With
    Exit Sub
EfectoParalisisUser_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoParalisisUser", Erl)
End Sub

Public Sub EfectoVelocidadUser(ByVal UserIndex As Integer)
    On Error GoTo EfectoVelocidadUser_Err
    If UserList(UserIndex).Counters.velocidad > 0 Then
        UserList(UserIndex).Counters.velocidad = UserList(UserIndex).Counters.velocidad - 1
    Else
        UserList(UserIndex).flags.VelocidadHechizada = 0
        Call ActualizarVelocidadDeUsuario(UserIndex)
    End If
    Exit Sub
EfectoVelocidadUser_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoVelocidadUser", Erl)
End Sub

Public Sub EfectoMaldicionUser(ByVal UserIndex As Integer)
    On Error GoTo EfectoMaldicionUser_Err
    If UserList(UserIndex).Counters.Maldicion > 0 Then
        UserList(UserIndex).Counters.Maldicion = UserList(UserIndex).Counters.Maldicion - 1
    Else
        UserList(UserIndex).flags.Maldicion = 0
        ' Msg518=¡La magia perdió su efecto! Ya puedes atacar.
        Call WriteLocaleMsg(UserIndex, 518, e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
    End If
    Exit Sub
EfectoMaldicionUser_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoMaldicionUser", Erl)
End Sub

Public Sub EfectoInmoUser(ByVal UserIndex As Integer)
    On Error GoTo EfectoInmoUser_Err
    With UserList(UserIndex)
        If .Counters.Inmovilizado > 0 Then
            .Counters.Inmovilizado = .Counters.Inmovilizado - 1
        Else
            .flags.Inmovilizado = 0
            If .clase = e_Class.Warrior Or .clase = e_Class.Hunter Or .clase = e_Class.Thief Or .clase = e_Class.Pirat Then
                .Counters.TiempoDeInmunidadParalisisNoMagicas = 3
            End If
            Call WriteInmovilizaOK(UserIndex)
        End If
    End With
    Exit Sub
EfectoInmoUser_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoInmoUser", Erl)
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
    On Error GoTo RecStamina_Err
    Dim trigger As Byte
    Dim Suerte  As Integer
    With UserList(UserIndex)
        trigger = MapData(.pos.Map, .pos.x, .pos.y).trigger
        If trigger = 1 And trigger = 2 And trigger = 4 Then Exit Sub
        If .Stats.MinSta < .Stats.MaxSta Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
                Exit Sub
            End If
            .Counters.STACounter = 0
            If .Counters.Trabajando > 0 Or IsSet(.flags.StatusMask, ePreventEnergyRestore) Then Exit Sub  'Trabajando no sube energía. (ToxicWaste)
            EnviarStats = True
            Select Case .Stats.UserSkills(e_Skill.Supervivencia)
                Case 0 To 10
                    Suerte = 5
                Case 11 To 20
                    Suerte = 7
                Case 21 To 30
                    Suerte = 9
                Case 31 To 40
                    Suerte = 11
                Case 41 To 50
                    Suerte = 13
                Case 51 To 60
                    Suerte = 15
                Case 61 To 70
                    Suerte = 17
                Case 71 To 80
                    Suerte = 19
                Case 81 To 90
                    Suerte = 21
                Case 91 To 99
                    Suerte = 23
                Case 100
                    Suerte = 25
            End Select
            Dim NuevaStamina As Long
            If .clase = e_Class.Trabajador Then
                NuevaStamina = .Stats.MinSta + RandomNumber(1, CInt(Porcentaje(.Stats.MaxSta, Suerte)))
            Else
                NuevaStamina = .Stats.MinSta + RandomNumber(1, CInt(Porcentaje(.Stats.MaxSta, Suerte)) / 1.6)
            End If
            ' Jopi: Prevenimos overflow al acotar la stamina que se puede recuperar en cualquier caso.
            ' Cuando te editabas la energia con el GM causaba este error.
            If NuevaStamina < 32000 Then
                .Stats.MinSta = NuevaStamina
            Else
                .Stats.MinSta = 32000
            End If
            If .Stats.MinSta > .Stats.MaxSta Then
                .Stats.MinSta = .Stats.MaxSta
            End If
        End If
    End With
    Exit Sub
RecStamina_Err:
    Call TraceError(Err.Number, Err.Description, "General.RecStamina", Erl)
End Sub

Public Sub PierdeEnergia(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
    On Error GoTo RecStamina_Err
    With UserList(UserIndex)
        If .Stats.MinSta > 0 Then
            If .Counters.STACounter < Intervalo Then
                .Counters.STACounter = .Counters.STACounter + 1
            Else
                .Counters.STACounter = 0
                EnviarStats = True
                Dim Cantidad As Integer
                Cantidad = RandomNumber(1, Porcentaje(.Stats.MaxSta, (MAXSKILLPOINTS * 1.5 - .Stats.UserSkills(e_Skill.Supervivencia)) * 0.25))
                .Stats.MinSta = .Stats.MinSta - Cantidad
                If .Stats.MinSta < 0 Then
                    .Stats.MinSta = 0
                End If
            End If
        End If
    End With
    Exit Sub
RecStamina_Err:
    Call TraceError(Err.Number, Err.Description, "General.PierdeEnergia", Erl)
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
    On Error GoTo EfectoVeneno_Err
    Dim Damage As Long
    If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
        UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
    Else
        Call CancelExit(UserIndex)
        With UserList(UserIndex)
            'Msg47=Estás envenenado, si no te curas morirás.
            Call WriteLocaleMsg(UserIndex, 47, e_FontTypeNames.FONTTYPE_VENENO)
            UserList(UserIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticleEffects.PoisonGas, 30, False, , UserList(UserIndex).pos.x, _
                    UserList(UserIndex).pos.y))
            .Counters.Veneno = 0
            ' El veneno saca un porcentaje de vida random.
            Damage = RandomNumber(3, 5)
            Damage = (1 + Damage * .Stats.MaxHp \ 100) ' Redondea para arriba
            If .ChatCombate = 1 Then
                ' "El veneno te ha causado ¬1 puntos de daño."
                Call WriteLocaleMsg(UserIndex, 390, e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(Damage))
            End If
            If UserMod.ModifyHealth(UserIndex, -Damage) Then
                Call CustomScenarios.UserDie(UserIndex)
                Call UserMod.UserDie(UserIndex)
            End If
        End With
    End If
    Exit Sub
EfectoVeneno_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoVeneno", Erl)
End Sub

' El incineramiento tiene una logica particular, que es hacer daño sostenido en el tiempo.
Public Sub EfectoIncineramiento(ByVal UserIndex As Integer)
    On Error GoTo EfectoIncineramiento_Err
    Dim Damage As Integer
    With UserList(UserIndex)
        ' 4 Mini intervalitos, dentro del intervalo total de incineracion
        If .Counters.Incineracion Mod (IntervaloIncineracion \ 4) = 0 Then
            ' "Te estás incinerando, si no te curas morirás.
            Call WriteLocaleMsg(UserIndex, 392, e_FontTypeNames.FONTTYPE_FIGHT)
            UserList(UserIndex).Counters.timeFx = 3
            Damage = RandomNumber(20, 30)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 73, 0, .pos.x, .pos.y))
            If .ChatCombate = 1 Then
                Call WriteLocaleMsg(UserIndex, 391, e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(Damage))
            End If
            If UserMod.ModifyHealth(UserIndex, -Damage) Then
                Call CustomScenarios.UserDie(UserIndex)
                Call UserMod.UserDie(UserIndex)
            End If
        End If
        .Counters.Incineracion = .Counters.Incineracion + 1
        If .Counters.Incineracion > IntervaloIncineracion Then
            ' Se termino la incineracion
            .flags.Incinerado = 0
            .Counters.Incineracion = 0
            Exit Sub
        End If
    End With
    Exit Sub
EfectoIncineramiento_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoIncineramiento", Erl)
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
    On Error GoTo DuracionPociones_Err
    'Controla la duracion de las pociones
    If UserList(UserIndex).flags.DuracionEfecto > 0 Then
        UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1
        If UserList(UserIndex).flags.DuracionEfecto = 0 Then
            UserList(UserIndex).flags.TomoPocion = False
            UserList(UserIndex).flags.TipoPocion = 0
            'volvemos los atributos al estado normal
            Dim LoopX As Integer
            For LoopX = 1 To NUMATRIBUTOS
                UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
            Next
            Call WriteFYA(UserIndex)
        End If
    End If
    Exit Sub
DuracionPociones_Err:
    Call TraceError(Err.Number, Err.Description, "General.DuracionPociones", Erl)
End Sub

Public Function HambreYSed(ByVal UserIndex As Integer) As Boolean
    On Error GoTo HambreYSed_Err
    If (UserList(UserIndex).flags.Privilegios And e_PlayerType.User) = 0 Then Exit Function
    'Sed
    If UserList(UserIndex).Stats.MinAGU > 0 Then
        If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
            UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
        Else
            UserList(UserIndex).Counters.AGUACounter = 0
            UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
            If UserList(UserIndex).Stats.MinAGU <= 0 Then
                UserList(UserIndex).Stats.MinAGU = 0
            End If
            HambreYSed = True
        End If
    End If
    'hambre
    If UserList(UserIndex).Stats.MinHam > 0 Then
        If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
            UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
        Else
            UserList(UserIndex).Counters.COMCounter = 0
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10
            If UserList(UserIndex).Stats.MinHam <= 0 Then
                UserList(UserIndex).Stats.MinHam = 0
            End If
            HambreYSed = True
        End If
    End If
    Exit Function
HambreYSed_Err:
    Call TraceError(Err.Number, Err.Description, "General.HambreYSed", Erl)
End Function

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
    On Error GoTo Sanar_Err
    ' Desnudo no regenera vida
    If UserList(UserIndex).flags.Desnudo = 1 Then Exit Sub
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).trigger = 1 And MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, _
            UserList(UserIndex).pos.y).trigger = 2 And MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).trigger = 4 Then Exit Sub
    Dim mashit As Integer
    'con el paso del tiempo va sanando....pero muy lentamente ;-)
    If UserList(UserIndex).flags.RegeneracionHP = 1 Then
        Intervalo = 400
    End If
    If UserList(UserIndex).Counters.HPCounter < Intervalo Then
        UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
    Else
        mashit = RandomNumber(Porcentaje(UserList(UserIndex).Stats.MaxHp, 5), Porcentaje(UserList(UserIndex).Stats.MaxHp, 10)) * UserMod.GetSelfHealingBonus(UserList(UserIndex))
        UserList(UserIndex).Counters.HPCounter = 0
        Call UserMod.ModifyHealth(UserIndex, mashit)
        ' Msg519=Has sanado.
        Call WriteLocaleMsg(UserIndex, 519, e_FontTypeNames.FONTTYPE_INFO)
        EnviarStats = True
    End If
    Exit Sub
Sanar_Err:
    Call TraceError(Err.Number, Err.Description, "General.Sanar", Erl)
End Sub

Public Sub CargaNpcsDat(Optional ByVal ActualizarNPCsExistentes As Boolean = False)
    On Error GoTo CargaNpcsDat_Err
    ' Leemos el NPCs.dat y lo almacenamos en la memoria.
    Set LeerNPCs = New clsIniManager
    Call LeerNPCs.Initialize(DatPath & "NPCs.dat")
    Call BuildNpcInfoCache
    ' Cargamos la lista de NPC's hostiles disponibles para spawnear.
    Call CargarSpawnList
    ' Actualizamos la informacion de los NPC's ya spawneados.
    If ActualizarNPCsExistentes Then
        Dim i As Long
        For i = 1 To NumNPCs
            If NpcList(i).flags.NPCActive Then
                Call OpenNPC(CInt(i), False, True)
            End If
            DoEvents
        Next i
    End If
    Exit Sub
CargaNpcsDat_Err:
    Call TraceError(Err.Number, Err.Description, "General.CargaNpcsDat", Erl)
End Sub

Sub PasarSegundo()
    On Error GoTo ErrHandler
    Dim i    As Long
    Dim h    As Byte
    Dim Mapa As Integer
    Dim x    As Byte
    Dim y    As Byte
    If TiempoPesca > 0 Then TiempoPesca = TiempoPesca + 1
    If CuentaRegresivaTimer > 0 Then
        If CuentaRegresivaTimer > 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1655, CuentaRegresivaTimer - 1, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1655=¬1 segundos...!
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1656, vbNullString, e_FontTypeNames.FONTTYPE_FIGHT)) 'Msg1656=¡Ya!!
        End If
        CuentaRegresivaTimer = CuentaRegresivaTimer - 1
    End If
    If Not InstanciaCaptura Is Nothing Then
        Call InstanciaCaptura.PasarSegundo
    End If
    segundos = segundos + 1
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                If .Counters.timeChat > 0 Then
                    .Counters.timeChat = .Counters.timeChat - 1
                End If
                If .Counters.LastTrabajo > 0 Then
                    .Counters.LastTrabajo = .Counters.LastTrabajo - 1
                End If
                If .Counters.timeFx > 0 Then
                    .Counters.timeFx = .Counters.timeFx - 1
                End If
                If .Counters.timeGuildChat > 0 Then
                    .Counters.timeGuildChat = .Counters.timeGuildChat - 1
                End If
                If .flags.Silenciado = 1 Then
                    .flags.SegundosPasados = .flags.SegundosPasados + 1
                    If .flags.SegundosPasados = 60 Then
                        .flags.MinutosRestantes = .flags.MinutosRestantes - 1
                        .flags.SegundosPasados = 0
                    End If
                    If .flags.MinutosRestantes = 0 Then
                        .flags.SegundosPasados = 0
                        .flags.Silenciado = 0
                        .flags.MinutosRestantes = 0
                        'Msg1018= Has sido liberado del silencio.
                        Call WriteLocaleMsg(i, "1018", e_FontTypeNames.FONTTYPE_SERVER)
                    End If
                End If
                If .flags.Muerto = 0 Then
                    Call DuracionPociones(i)
                    If .flags.invisible = 1 Or .Counters.DisabledInvisibility > 0 Then Call EfectoInvisibilidad(i)
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(i)
                    If .flags.Inmovilizado = 1 Then Call EfectoInmoUser(i)
                    If .flags.Ceguera = 1 Then Call EfectoCeguera(i)
                    If .flags.Estupidez = 1 Then Call EfectoEstupidez(i)
                    If .flags.Maldicion = 1 Then Call EfectoMaldicionUser(i)
                    If .flags.VelocidadHechizada > 0 Then Call EfectoVelocidadUser(i)
                    If HambreYSed(i) Then
                        Call WriteUpdateHungerAndThirst(i)
                    End If
                Else
                    If .flags.Traveling <> 0 Then Call TravelingEffect(i)
                End If
                If .Counters.TimerBarra > 0 Then
                    .Counters.TimerBarra = .Counters.TimerBarra - 1
                    If .Counters.TimerBarra = 0 Then
                        Select Case .Accion.TipoAccion
                            Case e_AccionBarra.Hogar
                                Call HomeArrival(i)
                            Case e_AccionBarra.Runa
                                Call CompletarAccionFin(i)
                        End Select
                        .Accion.Particula = 0
                        .Accion.TipoAccion = e_AccionBarra.CancelarAccion
                        .Accion.HechizoPendiente = 0
                        .Accion.RunaObj = 0
                        .Accion.ObjSlot = 0
                        .Accion.AccionPendiente = False
                    End If
                End If
                If .flags.UltimoMensaje > 0 Then
                    .Counters.RepetirMensaje = .Counters.RepetirMensaje + 1
                    If .Counters.RepetirMensaje >= 3 Then
                        .flags.UltimoMensaje = 0
                        .Counters.RepetirMensaje = 0
                    End If
                End If
                If .Counters.CuentaRegresiva >= 0 Then
                    If .Counters.CuentaRegresiva > 0 Then
                        Call WriteConsoleMsg(i, ">>>  " & .Counters.CuentaRegresiva & "  <<<", e_FontTypeNames.FONTTYPE_New_Gris)
                    Else
                        'Msg1019= >>> YA! <<<
                        Call WriteLocaleMsg(i, "1019", e_FontTypeNames.FONTTYPE_FIGHT)
                        Call WriteStopped(i, False)
                    End If
                    .Counters.CuentaRegresiva = .Counters.CuentaRegresiva - 1
                End If
                If .flags.Portal > 1 Then
                    .flags.Portal = .flags.Portal - 1
                    If .flags.Portal = 1 Then
                        Mapa = .flags.PortalM
                        x = .flags.PortalX
                        y = .flags.PortalY
                        Call SendData(SendTarget.toMap, .flags.PortalM, PrepareMessageParticleFXToFloor(x, y, e_GraphicEffects.TpVerde, 0))
                        Call SendData(SendTarget.toMap, .flags.PortalM, PrepareMessageLightFXToFloor(x, y, 0, 105))
                        If MapData(Mapa, x, y).TileExit.Map > 0 Then
                            MapData(Mapa, x, y).TileExit.Map = 0
                            MapData(Mapa, x, y).TileExit.x = 0
                            MapData(Mapa, x, y).TileExit.y = 0
                        End If
                        MapData(Mapa, x, y).Particula = 0
                        MapData(Mapa, x, y).TimeParticula = 0
                        MapData(Mapa, x, y).Particula = 0
                        MapData(Mapa, x, y).TimeParticula = 0
                        .flags.Portal = 0
                        .flags.PortalM = 0
                        .flags.PortalY = 0
                        .flags.PortalX = 0
                        .flags.PortalMDestino = 0
                        .flags.PortalYDestino = 0
                        .flags.PortalXDestino = 0
                    End If
                End If
                If .Counters.EnCombate > 0 Then
                    .Counters.EnCombate = .Counters.EnCombate - 1
                End If
                If .Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
                    .Counters.TiempoDeInmunidadParalisisNoMagicas = .Counters.TiempoDeInmunidadParalisisNoMagicas - 1
                End If
                If .Counters.TiempoDeInmunidad > 0 Then
                    .Counters.TiempoDeInmunidad = .Counters.TiempoDeInmunidad - 1
                    If .Counters.TiempoDeInmunidad = 0 Then
                        .flags.Inmunidad = 0
                    End If
                End If
                If .flags.Subastando Then
                    .Counters.TiempoParaSubastar = .Counters.TiempoParaSubastar - 1
                    If .Counters.TiempoParaSubastar = 0 Then
                        Call CancelarSubasta
                    End If
                End If
                'Cerrar usuario
                If .Counters.Saliendo Then
                    '  If .flags.Muerto = 1 Then .Counters.Salir = 0
                    .Counters.Salir = .Counters.Salir - 1
                    ' Call WriteConsoleMsg(i, "Se saldrá del juego en " & .Counters.Salir & " segundos...", e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(i, "203", e_FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
                    If .Counters.Salir <= 0 Then
                        'Msg1020= Gracias por jugar Argentum 20.
                        Call WriteLocaleMsg(i, "1020", e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(i)
                        Call CloseSocket(i)
                    End If
                End If
            End If ' If UserLogged
        End With
    Next i
    ' **********************************
    ' **********  Invasiones  **********
    ' **********************************
    For i = 1 To UBound(Invasiones)
        With Invasiones(i)
            ' Si la invasión está activa
            If .Activa Then
                .TimerSpawn = .TimerSpawn + 1
                ' Comprobamos si hay que spawnear NPCs
                If .TimerSpawn >= .IntervaloSpawn Then
                    Call InvasionSpawnNPC(i)
                    .TimerSpawn = 0
                End If
                ' ------------------------------------
                .TimerMostrarInfo = .TimerMostrarInfo + 1
                ' Comprobamos si hay que mostrar la info
                If .TimerMostrarInfo >= 5 Then
                    Call EnviarInfoInvasion(i)
                    .TimerMostrarInfo = 0
                End If
            End If
        End With
    Next
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "General.PasarSegundo", Erl)
End Sub

Sub GuardarUsuarios()
    On Error GoTo GuardarUsuarios_Err
    haciendoBK = True
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1657, vbNullString, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1657=Servidor » Grabando Personajes
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call modNetwork.Poll
        End If
    Next i
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i)
        End If
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1658, vbNullString, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1658=Servidor » Personajes Grabados
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    haciendoBK = False
    Exit Sub
GuardarUsuarios_Err:
    Call TraceError(Err.Number, Err.Description, "General.GuardarUsuarios", Erl)
End Sub

Public Sub FreeNPCs()
    On Error GoTo FreeNPCs_Err
    'Releases all NPC Indexes
    Dim LoopC As Long
    ' Free all NPC indexes
    For LoopC = 1 To MaxNPCs
        Call ReleaseNpc(LoopC, e_DeleteSource.eReleaseAll)
    Next LoopC
    Exit Sub
FreeNPCs_Err:
    Call TraceError(Err.Number, Err.Description, "General.FreeNPCs", Erl)
End Sub

Public Sub FreeCharIndexes()
    'Releases all char indexes
    ' Free all char indexes (set them all to 0)
    On Error GoTo FreeCharIndexes_Err
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
    Exit Sub
FreeCharIndexes_Err:
    Call TraceError(Err.Number, Err.Description, "General.FreeCharIndexes", Erl)
End Sub

Function RandomString(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
    On Error GoTo RandomString_Err
    Randomize Time
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    If OnlyUpper Then
        rgch = UCase(rgch)
    Else
        rgch = rgch & UCase(rgch)
    End If
    rgch = rgch & "0123456789"  ' & "#@!~$()-_"
    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next
    Exit Function
RandomString_Err:
    Call TraceError(Err.Number, Err.Description, "General.RandomString", Erl)
End Function

Function RandomName(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
    On Error GoTo RandomString_Err
    Randomize Time
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    If OnlyUpper Then
        rgch = UCase$(rgch)
    Else
        rgch = rgch & UCase$(rgch)
    End If
    Dim i As Long
    For i = 1 To cb
        RandomName = RandomName & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next
    Exit Function
RandomString_Err:
    Call TraceError(Err.Number, Err.Description, "General.RandomString", Erl)
End Function

'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
    On Error GoTo errHnd
    Dim lPos As Long
    Dim lX   As Long
    Dim iAsc As Integer
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then Exit Function
            End If
        Next lX
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    On Error GoTo CMSValidateChar__Err
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
    Exit Function
CMSValidateChar__Err:
    Call TraceError(Err.Number, Err.Description, "General.CMSValidateChar_", Erl)
End Function

Public Function Tilde(ByRef data As String) As String
    On Error GoTo Tilde_Err
    Tilde = UCase$(data)
    Tilde = Replace$(Tilde, "Á", "A")
    Tilde = Replace$(Tilde, "É", "E")
    Tilde = Replace$(Tilde, "Í", "I")
    Tilde = Replace$(Tilde, "Ó", "O")
    Tilde = Replace$(Tilde, "Ú", "U")
    Exit Function
Tilde_Err:
    Call TraceError(Err.Number, Err.Description, "Mod_General.Tilde", Erl)
End Function

Public Sub CerrarServidor()
    'Save stats!!!
    Call frmMain.QuitarIconoSystray
    ' Limpieza del socket del servidor.
    Call modNetwork.Disconnect
    Dim LoopC As Long
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).ConnectionDetails.ConnIDValida Then
            Call CloseSocket(LoopC)
        End If
    Next
    Call UnloadAntiCheat
    If Database_Enabled Then Database_Close
    End
End Sub

Public Function PonerPuntos(ByVal Numero As Long) As String
    On Error GoTo PonerPuntos_Err
    Dim i     As Integer
    Dim Cifra As String
    Cifra = str$(Numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)
    For i = 0 To 4
        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
            End If
        Else
            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
            End If
            Exit For
        End If
    Next
    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
    Exit Function
PonerPuntos_Err:
    Call TraceError(Err.Number, Err.Description, "ModLadder.PonerPuntos", Erl)
End Function

' Autor: WyroX
Function CalcularPromedioVida(ByVal UserIndex As Integer) As Double
    With UserList(UserIndex)
        If .Stats.ELV = 1 Then
            ' Siempre estamos promedio al lvl 1
            CalcularPromedioVida = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
        Else
            CalcularPromedioVida = (.Stats.MaxHp - .Stats.UserAtributos(e_Atributos.Constitucion)) / (.Stats.ELV - 1)
        End If
    End With
End Function

' Adaptado desde https://stackoverflow.com/questions/29325069/how-to-generate-random-numbers-biased-towards-one-value-in-a-range/29325222#29325222
' By WyroX
Function RandomIntBiased(ByVal Min As Double, ByVal max As Double, ByVal Bias As Double, ByVal Influence As Double) As Double
    On Error GoTo handle
    Dim RandomRango As Double, Mix As Double
    ' Rnd: número pseudo-aleatorio entre 0 y 1
    ' RandomRango: Nuevo aumento de vida
    RandomRango = Rnd * (max - Min) + Min
    ' Mix: Qué tanto afectamos a la vida random que salió con el sesgo Bias
    ' El bias hace tender el promedio actual del personaje al promedio de manual
    Mix = Rnd * Influence
    ' RandomIntBiased: Valor final de vida
    ' Ejemplo:
    ' Si Mix=0.1, 10% de influencia del Bias sobre el valor final de vida
    ' RandomIntBiased = RandomRango 0.9 + Bias 0.1
    RandomIntBiased = RandomRango * (1 - Mix) + Bias * Mix
    Exit Function
handle:
    Call TraceError(Err.Number, Err.Description, "General.RandomIntBiased")
    RandomIntBiased = Bias
End Function

'Very efficient function for testing whether this code is running in the IDE or compiled
'https://www.vbforums.com/showthread.php?231468-VB-Detect-if-you-are-running-in-the-IDE&p=5413357&viewfull=1#post5413357
Public Function RunningInVB(Optional ByRef b As Boolean = True) As Boolean
    If b Then Debug.Assert Not RunningInVB(RunningInVB) Else b = True
End Function

'  Mensaje a todo el mundo
Public Sub MensajeGlobal(texto As String, Fuente As e_FontTypeNames)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(texto, Fuente))
End Sub

'  Devuelve si X e Y están dentro del Rectangle
Public Function InsideRectangle(r As t_Rectangle, ByVal x As Integer, ByVal y As Integer) As Boolean
    If x < r.X1 Then Exit Function
    If x > r.X2 Then Exit Function
    If y < r.Y1 Then Exit Function
    If y > r.Y2 Then Exit Function
    InsideRectangle = True
End Function

' Based on: https://stackoverflow.com/questions/1378604/end-process-from-task-manager-using-vb-6-code (ultima respuesta)
Public Function GetProcessCount(ByVal processName As String) As Byte
    Dim oService     As Object
    Dim servicename  As String
    Dim processCount As Byte
    Dim oWMI         As Object: Set oWMI = GetObject("winmgmts:")
    Dim oServices    As Object: Set oServices = oWMI.InstancesOf("win32_process")
    For Each oService In oServices
        servicename = CStr(oService.name)
        If StrComp(servicename, processName, vbTextCompare) = 0 Then
            ' Para matar un proceso adentro de este loop usar.
            ' oService.Terminate
            processCount = processCount + 1
        End If
    Next
    GetProcessCount = processCount
End Function

Public Function EsMapaInterdimensional(ByVal Map As Integer) As Boolean
    Dim i As Integer
    For i = 1 To UBound(MapasInterdimensionales)
        If Map = MapasInterdimensionales(i) Then
            EsMapaInterdimensional = True
            Exit Function
        End If
    Next
End Function

Public Function IsValidIPAddress(ByVal IP As String) As Boolean
    On Error GoTo Handler
    Dim varAddress As Variant, n As Long, lCount As Long
    varAddress = Split(IP, ".", 4, vbTextCompare)
    If IsArray(varAddress) Then
        For n = LBound(varAddress) To UBound(varAddress)
            lCount = lCount + 1
            varAddress(n) = CByte(varAddress(n))
        Next
        IsValidIPAddress = (lCount = 4)
    End If
Handler:
End Function

Function Ceil(x As Variant) As Variant
    On Error GoTo Ceil_Err
    Ceil = IIf(Fix(x) = x, x, Fix(x) + 1)
    Exit Function
Ceil_Err:
    Call TraceError(Err.Number, Err.Description & "Ceil_Err", Erl)
End Function

Function Clamp(x As Variant, a As Variant, b As Variant) As Variant
    On Error GoTo Clamp_Err
    Clamp = IIf(x < a, a, IIf(x > b, b, x))
    Exit Function
Clamp_Err:
    Call TraceError(Err.Number, Err.Description & "Clamp_Err", Erl)
End Function

Private Function GetElapsed() As Single
    Static sTime1     As Currency
    Static sTime2     As Currency
    Static sFrequency As Currency
    'Get the timer frequency
    If sFrequency = 0 Then
        Call QueryPerformanceFrequency(sFrequency)
    End If
    'Get current time
    Call QueryPerformanceCounter(sTime1)
    'Calculate elapsed time
    GetElapsed = ((sTime1 - sTime2) / sFrequency * 1000)
    'Get next end time
    Call QueryPerformanceCounter(sTime2)
End Function

Public Function RunScriptInFile(ByVal FilePath As String) As Boolean
    Dim script As String
    script = FileText(FilePath)
    script = Replace(Replace(script, Chr(10), ""), Chr(13), "")
    Dim RS As Recordset
    If script <> vbNullString Then
        Set RS = Query(script)
        If RS Is Nothing Then
            RunScriptInFile = False
            Exit Function
        End If
    End If
    RunScriptInFile = True
End Function

'Reads the files inside the ScriptsDB folder, it can be a create table, alter, etc.
'we are calling this files dbmigrations, this function check this
'folder and the db, and run all the files that are not registered in the db migration table
'the file should store the name in the format of YYYYMMDD-XX-description text.sql
'where the XX is the number of migrations generated the same day
Public Sub LoadDBMigrations()
    On Error GoTo LoadDBMigrations_Err
    'Consulto a la DB a ver si existe la tabla migrations
    Dim RS As Recordset
    Set RS = Query("select * from migrations")
    Dim LastScript As String: LastScript = ""
    If RS Is Nothing Then
        Call Query("CREATE TABLE ""migrations"" (    ""id"" INTEGER NOT NULL,    ""date"" VARCHAR(11) NOT NULL,    ""description"" VARCHAR(50) NULL,    Primary key(""id""));")
    Else
        Set RS = Query("select date from migrations order by id desc LIMIT 1;")
        If RS.RecordCount > 0 Then LastScript = RS!Date
    End If
    Dim sFilename As String
    sFilename = dir(App.Path & "/ScriptsDB/")
    Do While sFilename <> ""
        If Len(sFilename) > 11 Then
            Dim date_ As String
            date_ = Left(sFilename, 11)
            If LastScript < date_ Then
                'Leemos el archivo
                Dim script      As String
                Dim Description As String
                Description = mid(sFilename, 13, Len(sFilename) - 16)
                If RunScriptInFile(App.Path & "/ScriptsDB/" & sFilename) Then
                    Call Query("insert into migrations (date, description) values (?,?);", date_, Description)
                Else
                    Call Err.raise(5, , "invalid - " & Description)
                End If
            End If
        End If
        sFilename = dir()
    Loop
    Exit Sub
LoadDBMigrations_Err:
    Call TraceError(Err.Number, Err.Description, "modGuilds.LoadDBMigrations", Erl)
    Call MsgBox(DBError & vbNewLine & "Script:" & Err.Description, vbCritical, "ERROR MIGRATIONS")
End Sub

Function FileText(Filename$) As String
    Dim handle As Integer
    handle = FreeFile
    Open Filename$ For Input As #handle
    FileText = Input$(LOF(handle), handle)
    Close #handle
End Function

Public Function IsArrayInitialized(ByRef arr) As Boolean
    Dim rv As Long
    On Error Resume Next
    rv = UBound(arr)
    IsArrayInitialized = (Err.Number = 0) And rv >= 0
End Function
