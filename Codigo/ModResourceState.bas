' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
Option Explicit

' Lleva registro de los tiles de yacimientos/arboles que estan agotados o
' parcialmente consumidos, para no tener que escanear todo el mapa al guardar.
' Key: "Map|x|y", Value: sin uso (marcador).
Public ResourceDirtyTiles As New Dictionary

Private Const RESOURCE_STATE_DIR  As String = "SourcesStates\"
Private Const RESOURCE_STATE_FILE As String = "RecursosEstado.dat"

Public Sub EnsureResourceStateDir()
    On Error GoTo EnsureResourceStateDir_Err
    If Len(dir$(App.path & "\" & RESOURCE_STATE_DIR, vbDirectory)) = 0 Then
        MkDir App.path & "\" & RESOURCE_STATE_DIR
    End If
    Exit Sub
EnsureResourceStateDir_Err:
    Call TraceError(Err.Number, Err.Description, "ModResourceState.EnsureResourceStateDir", Erl)
End Sub

Public Sub MarkResourceDirty(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo MarkResourceDirty_Err
    Dim key As String
    key = Map & "|" & x & "|" & y
    If Not ResourceDirtyTiles.Exists(key) Then
        Call ResourceDirtyTiles.Add(key, True)
    End If
    Exit Sub
MarkResourceDirty_Err:
    Call TraceError(Err.Number, Err.Description, "ModResourceState.MarkResourceDirty", Erl)
End Sub

Public Sub ClearResourceDirty(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte)
    On Error GoTo ClearResourceDirty_Err
    Dim key As String
    key = Map & "|" & x & "|" & y
    If ResourceDirtyTiles.Exists(key) Then
        Call ResourceDirtyTiles.Remove(key)
    End If
    Exit Sub
ClearResourceDirty_Err:
    Call TraceError(Err.Number, Err.Description, "ModResourceState.ClearResourceDirty", Erl)
End Sub

' Guarda unicamente los tiles marcados como "dirty" (agotados o parcialmente consumidos).
' Reescribe el archivo entero cada vez para no arrastrar entradas viejas que ya se regeneraron.
Public Sub SaveResourceState()
    On Error GoTo SaveResourceState_Err
    Call EnsureResourceStateDir
    Dim n    As Integer
    Dim key  As Variant
    Dim Parts() As String
    Dim Map  As Integer
    Dim x    As Byte
    Dim y    As Byte
    n = FreeFile
    Open App.path & "\" & RESOURCE_STATE_DIR & RESOURCE_STATE_FILE For Output As #n
    Print #n, ResourceDirtyTiles.count
    For Each key In ResourceDirtyTiles.Keys
        Parts = Split(key, "|")
        Map = CInt(Parts(0))
        x = CByte(Parts(1))
        y = CByte(Parts(2))
        Print #n, Map & "|" & x & "|" & y & "|" & MapData(Map, x, y).ObjInfo.amount & "|" & MapData(Map, x, y).ResourceLastUseEpoch
    Next key
    Close #n
    Exit Sub
SaveResourceState_Err:
    Call TraceError(Err.Number, Err.Description, "ModResourceState.SaveResourceState", Erl)
End Sub

' Carga el estado persistido de recursos. Debe llamarse DESPUES de LoadMapData/CargarBackUp,
' ya que pisa el amount/data que esas funciones dejaron en full para los tiles que corresponda.
Public Sub LoadResourceState()
    On Error GoTo LoadResourceState_Err
    Dim FilePath As String
    FilePath = App.path & "\" & RESOURCE_STATE_DIR & RESOURCE_STATE_FILE
    If Not FileExist(FilePath) Then Exit Sub
    Dim n        As Integer
    Dim count    As Long
    Dim i        As Long
    Dim TextLine As String
    Dim Parts()  As String
    n = FreeFile
    Open FilePath For Input As #n
    If EOF(n) Then
        Close n
        Exit Sub
    End If
    Line Input #n, TextLine
    count = val(TextLine)
    Dim nowEpoch As Long
    nowEpoch = CLng(DateDiff("s", "01/01/1970", Now))
    For i = 1 To count
        If EOF(n) Then Exit For
        Line Input #n, TextLine
        Parts = Split(TextLine, "|")
        If UBound(Parts) = 4 Then
            Dim Map         As Integer
            Dim x           As Byte
            Dim y           As Byte
            Dim SavedAmount As Integer
            Dim SavedEpoch  As Long
            Map = val(Parts(0))
            x = val(Parts(1))
            y = val(Parts(2))
            SavedAmount = val(Parts(3))
            SavedEpoch = val(Parts(4))
            If Map >= 1 And Map <= NumMaps Then
                Dim ObjIndex As Integer
                ObjIndex = MapData(Map, x, y).ObjInfo.ObjIndex
                If ObjIndex > 0 Then
                    If ObjData(ObjIndex).OBJType = e_OBJType.otOreDeposit Or ObjData(ObjIndex).OBJType = e_OBJType.otTrees Then
                        Dim elapsedRealSec As Long
                        elapsedRealSec = nowEpoch - SavedEpoch
                        If elapsedRealSec > ObjData(ObjIndex).TiempoRegenerar Then
                            ' Ya se regenero durante el tiempo que el server estuvo caido/parcheando.
                            MapData(Map, x, y).ObjInfo.amount = ObjData(ObjIndex).VidaUtil
                            MapData(Map, x, y).ObjInfo.data = &H7FFFFFFF
                            MapData(Map, x, y).ResourceLastUseEpoch = 0
                        Else
                            ' Todavia no se regenero: restauramos el amount real y reconstruimos
                            ' un tick sintetico para que ActualizarRecurso siga funcionando sin cambios.
                            MapData(Map, x, y).ObjInfo.amount = SavedAmount
                            Dim elapsedMsForTick As Double
                            elapsedMsForTick = CDbl(elapsedRealSec) * 1000#
                            MapData(Map, x, y).ObjInfo.data = AddMod32(GetTickCountRaw(), CLng(elapsedMsForTick Mod CDbl(TICKS32)) * -1)
                            MapData(Map, x, y).ResourceLastUseEpoch = SavedEpoch
                            Call MarkResourceDirty(Map, x, y)
                        End If
                    End If
                End If
            End If
        End If
    Next i
    Close n
    Exit Sub
LoadResourceState_Err:
    Call TraceError(Err.Number, Err.Description, "ModResourceState.LoadResourceState", Erl)
End Sub
