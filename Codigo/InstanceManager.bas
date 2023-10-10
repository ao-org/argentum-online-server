Attribute VB_Name = "InstanceManager"
Option Explicit


Private AvaibleInstanceMap As t_IndexHeap

Public Sub InitializeInstanceHeap(ByVal Size As Integer, ByVal MapIndexStart As Integer)
On Error GoTo ErrHandler_InitializeInstanceHeap
    ReDim AvaibleInstanceMap.IndexInfo(Size)
    Dim i As Integer
    For i = 1 To Size
        AvaibleInstanceMap.IndexInfo(i) = Size - (i - 1) + MapIndexStart
    Next i
    AvaibleInstanceMap.currentIndex = Size
    Exit Sub
ErrHandler_InitializeInstanceHeap:
    Call TraceError(Err.Number, Err.Description, "InstanceManager.InitializeInstanceHeap", Erl)
End Sub

Public Function ReleaseInstance(ByVal InstanceMapIndex As Integer) As Boolean
On Error GoTo ErrHandler
    AvaibleInstanceMap.currentIndex = AvaibleInstanceMap.currentIndex + 1
    Debug.Assert AvaibleInstanceMap.currentIndex <= UBound(AvaibleInstanceMap.IndexInfo)
    AvaibleInstanceMap.IndexInfo(AvaibleInstanceMap.currentIndex) = InstanceMapIndex
    ReleaseInstance = True
    Exit Function
ErrHandler:
    ReleaseInstance = False
    Call TraceError(Err.Number, Err.Description, "InstanceManager.ReleaseInstance", Erl)
End Function

Public Function GetAvailableInstanceCount() As Integer
    GetAvailableInstanceCount = AvaibleInstanceMap.currentIndex
End Function

Public Function GetNextAvailableInstance() As Integer
On Error GoTo ErrHandler
    If (AvaibleInstanceMap.currentIndex = 0) Then
        GetNextAvailableInstance = -1
        Return
    End If
    GetNextAvailableInstance = AvaibleInstanceMap.IndexInfo(AvaibleInstanceMap.currentIndex)
    AvaibleInstanceMap.currentIndex = AvaibleInstanceMap.currentIndex - 1
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "InstanceManager.GetNextAvailableInstance", Erl)
End Function

Public Sub CloneMap(ByVal SourceMapIndex As Integer, ByVal DestMapIndex As Integer)
    MapInfo(DestMapIndex) = MapInfo(SourceMapIndex)
    Dim PosX As Integer
    Dim PosY As Integer
    Dim Time As Long
    Time = GetTickCount()
    For PosY = YMinMapSize To YMaxMapSize
        For PosX = XMinMapSize To XMaxMapSize
            MapData(DestMapIndex, PosX, PosY) = MapData(SourceMapIndex, PosX, PosY)
        Next PosX
    Next PosY
    Time = GetTickCount() - Time
    
End Sub
