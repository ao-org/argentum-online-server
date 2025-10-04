Attribute VB_Name = "InstanceManager"
Option Explicit
Private AvailableInstanceMap As t_IndexHeap

Public Type t_TranslationMapping
    OriginalTarget As Integer
    NewTarget As Integer
End Type

Public Sub InitializeInstanceHeap(ByVal Size As Integer, ByVal MapIndexStart As Integer)
    On Error GoTo ErrHandler_InitializeInstanceHeap
    ReDim AvailableInstanceMap.IndexInfo(Size)
    Dim i As Integer
    For i = 1 To Size
        AvailableInstanceMap.IndexInfo(i) = Size - (i - 1) + MapIndexStart
    Next i
    AvailableInstanceMap.currentIndex = Size
    Exit Sub
ErrHandler_InitializeInstanceHeap:
    Call TraceError(Err.Number, Err.Description, "InstanceManager.InitializeInstanceHeap", Erl)
End Sub

Public Function ReleaseInstance(ByVal InstanceMapIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    AvailableInstanceMap.currentIndex = AvailableInstanceMap.currentIndex + 1
    Debug.Assert AvailableInstanceMap.currentIndex <= UBound(AvailableInstanceMap.IndexInfo)
    AvailableInstanceMap.IndexInfo(AvailableInstanceMap.currentIndex) = InstanceMapIndex
    ReleaseInstance = True
    MapInfo(InstanceMapIndex).MapResource = 0
    Exit Function
ErrHandler:
    ReleaseInstance = False
    Call TraceError(Err.Number, Err.Description, "InstanceManager.ReleaseInstance", Erl)
End Function

Public Function GetNextAvailableInstance() As Integer
    On Error GoTo ErrHandler
    If (AvailableInstanceMap.currentIndex = 0) Then
        GetNextAvailableInstance = -1
        Exit Function
    End If
    GetNextAvailableInstance = AvailableInstanceMap.IndexInfo(AvailableInstanceMap.currentIndex)
    AvailableInstanceMap.currentIndex = AvailableInstanceMap.currentIndex - 1
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "InstanceManager.GetNextAvailableInstance", Erl)
End Function

Public Sub CloneMap(ByVal SourceMapIndex As Integer, ByVal DestMapIndex As Integer)
    Dim Translations(0) As t_TranslationMapping
    Call CloneMapWithTranslations(SourceMapIndex, DestMapIndex, Translations)
End Sub

Public Sub CloneMapWithTranslations(ByVal SourceMapIndex As Integer, ByVal DestMapIndex As Integer, ByRef TranslationMappings() As t_TranslationMapping)
    MapInfo(DestMapIndex) = MapInfo(SourceMapIndex)
    MapInfo(DestMapIndex).MapResource = SourceMapIndex
    Dim PosX             As Integer
    Dim PosY             As Integer
    Dim PerformanceTimer As Long
    Dim i                As Integer
    Call PerformanceTestStart(PerformanceTimer)
    For PosY = YMinMapSize To YMaxMapSize
        For PosX = XMinMapSize To XMaxMapSize
            MapData(DestMapIndex, PosX, PosY) = MapData(SourceMapIndex, PosX, PosY)
            If (MapData(DestMapIndex, PosX, PosY).TileExit.Map > 0) Then
                For i = LBound(TranslationMappings) To UBound(TranslationMappings)
                    If MapData(DestMapIndex, PosX, PosY).TileExit.Map = TranslationMappings(i).OriginalTarget Then
                        MapData(DestMapIndex, PosX, PosY).TileExit.Map = TranslationMappings(i).NewTarget
                    End If
                Next i
            End If
        Next PosX
    Next PosY
    Call PerformTimeLimitCheck(PerformanceTimer, "CloneMapWithTranslations time", 50)
End Sub
