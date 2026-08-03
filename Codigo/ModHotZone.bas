Option Explicit

Private Type t_HotZoneEntry
    Name As String
    MapCount As Integer
    Maps() As Integer
End Type

Private HotZones() As t_HotZoneEntry
Private HotZoneCount As Integer
Private HotZoneLoaded As Boolean

Private HotZoneActive As Boolean
Private HotZoneActiveDate As Date
Private HotZoneCurrentZoneIndex As Integer
Private HotZoneReminderCounter As Integer

Public Sub LoadHotZones(Optional ByVal Filename As String = "HotZones.dat")
    On Error GoTo ErrHandler
    Dim reader As clsIniManager
    Set reader = New clsIniManager
    Call reader.Initialize(Filename)

    Dim RawCount As Integer
    RawCount = CInt(val(reader.GetValue("ZONES", "Count", 0)))
    If RawCount <= 0 Then
        HotZoneLoaded = False
        Call LogError("ModHotZone.LoadHotZones: HotZones.dat sin zonas configuradas (Count=0)")
        Exit Sub
    End If

    ReDim HotZones(1 To RawCount)
    Dim ValidCount As Integer
    ValidCount = 0

    Dim z As Integer
    For z = 1 To RawCount
        Dim SectionName As String
        SectionName = "ZONE" & z
        Dim ZoneName As String
        ZoneName = reader.GetValue(SectionName, "Name", vbNullString)
        Dim ZoneMapCount As Integer
        ZoneMapCount = CInt(val(reader.GetValue(SectionName, "MapCount", 0)))

        If LenB(ZoneName) = 0 Or ZoneMapCount <= 0 Then
            Call LogError("ModHotZone.LoadHotZones: " & SectionName & " invalida (falta Name o MapCount) en HotZones.dat")
        Else
            Dim TempMaps() As Integer
            ReDim TempMaps(1 To ZoneMapCount)
            Dim ValidMapCount As Integer
            ValidMapCount = 0
            Dim m As Integer
            For m = 1 To ZoneMapCount
                Dim MapNum As Integer
                MapNum = CInt(val(reader.GetValue(SectionName, "Map" & m, 0)))
                If MapNum > 0 Then
                    ValidMapCount = ValidMapCount + 1
                    TempMaps(ValidMapCount) = MapNum
                Else
                    Call LogError("ModHotZone.LoadHotZones: " & SectionName & ".Map" & m & " invalido o faltante en HotZones.dat")
                End If
            Next m

            If ValidMapCount > 0 Then
                ValidCount = ValidCount + 1
                HotZones(ValidCount).Name = ZoneName
                HotZones(ValidCount).MapCount = ValidMapCount
                ReDim HotZones(ValidCount).Maps(1 To ValidMapCount)
                Dim k As Integer
                For k = 1 To ValidMapCount
                    HotZones(ValidCount).Maps(k) = TempMaps(k)
                Next k
            Else
                Call LogError("ModHotZone.LoadHotZones: " & SectionName & " descartada, sin mapas validos")
            End If
        End If
    Next z

    HotZoneCount = ValidCount
    HotZoneLoaded = (HotZoneCount > 0)
    Set reader = Nothing
    Exit Sub
ErrHandler:
    HotZoneLoaded = False
    Call TraceError(Err.Number, Err.Description, "ModHotZone.LoadHotZones", Erl)
End Sub

Public Sub CheckHotZoneEvent()
    On Error GoTo ErrHandler
    If Not HotZoneLoaded Then Exit Sub

    Dim CurrentHour As Byte
    Dim StartHour   As Byte
    Dim EndHour     As Byte
    Dim EffectiveEndHour As Integer
    CurrentHour = Hour(Time)
    StartHour = SvrConfig.GetValue("HotZoneStartHour")
    EndHour = SvrConfig.GetValue("HotZoneEndHour")

    ' EndHour=0 significa "hasta medianoche" (fin del dia), no "antes de las 00hs"
    If EndHour = 0 Then
        EffectiveEndHour = 24
    Else
        EffectiveEndHour = EndHour
    End If

    If CurrentHour >= StartHour And CurrentHour < EffectiveEndHour Then
        If Not HotZoneActive Then
            If HotZoneActiveDate <> Date Then
                HotZoneCurrentZoneIndex = PickRandomHotZoneZone()
                HotZoneActiveDate = Date
            End If
            If HotZoneCurrentZoneIndex > 0 Then
                HotZoneActive = True
                HotZoneReminderCounter = 0
                Call AnnounceHotZoneStart(HotZoneCurrentZoneIndex, EndHour)
            End If
        Else
            HotZoneReminderCounter = HotZoneReminderCounter + 1
            If HotZoneReminderCounter >= 15 Then
                HotZoneReminderCounter = 0
                Call AnnounceHotZoneReminder
            End If
        End If
    Else
        If HotZoneActive Then
            HotZoneActive = False
            Call AnnounceHotZoneEnd
        End If
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModHotZone.CheckHotZoneEvent", Erl)
End Sub

Private Sub AnnounceHotZoneReminder()
    On Error GoTo ErrHandler
    If HotZoneCurrentZoneIndex <= 0 Then Exit Sub
    Dim ZoneName As String
    ZoneName = HotZones(HotZoneCurrentZoneIndex).Name
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_HOTZONE_REMINDER, ZoneName, e_FontTypeNames.FONTTYPE_DIOS))
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModHotZone.AnnounceHotZoneReminder", Erl)
End Sub

Private Function PickRandomHotZoneZone() As Integer
    If HotZoneCount <= 0 Then
        PickRandomHotZoneZone = 0
        Exit Function
    End If
    PickRandomHotZoneZone = RandomNumber(1, HotZoneCount)
End Function

Private Function BuildMapListString(ByRef Zone As t_HotZoneEntry) As String
    Dim Result As String
    Dim i As Integer
    For i = 1 To Zone.MapCount
        If i > 1 Then Result = Result & ", "
        Result = Result & CStr(Zone.Maps(i))
    Next i
    BuildMapListString = Result
End Function

Private Sub AnnounceHotZoneStart(ByVal ZoneIndex As Integer, ByVal EndHour As Byte)
    On Error GoTo ErrHandler
    Dim ZoneName As String
    ZoneName = HotZones(ZoneIndex).Name

    Dim ExtraParams As String
    ExtraParams = ZoneName & Chr$(172) & CStr(EndHour)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_HOTZONE_STARTED, ExtraParams, e_FontTypeNames.FONTTYPE_DIOS))
    Call AgregarAConsola("Servidor - Hot Zone activada: " & ZoneName & " hasta las " & EndHour & "hs.")
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModHotZone.AnnounceHotZoneStart", Erl)
End Sub

Private Sub AnnounceHotZoneEnd()
    On Error GoTo ErrHandler
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_HOTZONE_ENDED, vbNullString, e_FontTypeNames.FONTTYPE_DIOS))
    Call AgregarAConsola("Servidor - Hot Zone finalizada por hoy.")
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModHotZone.AnnounceHotZoneEnd", Erl)
End Sub

Private Function MapIsInCurrentZone(ByVal Map As Integer) As Boolean
    MapIsInCurrentZone = False
    If HotZoneCurrentZoneIndex <= 0 Then Exit Function
    Dim i As Integer
    For i = 1 To HotZones(HotZoneCurrentZoneIndex).MapCount
        If HotZones(HotZoneCurrentZoneIndex).Maps(i) = Map Then
            MapIsInCurrentZone = True
            Exit Function
        End If
    Next i
End Function

' Usado por GetExpForUser y CalcularDarExpGrupal.
' Devuelve 1 (neutro) si el evento no esta activo, si el mapa no pertenece a la zona activa, o ante cualquier valor de config invalido.
Public Function GetHotZoneExpMultiplier(ByVal Map As Integer) As Single
    If HotZoneActive And MapIsInCurrentZone(Map) Then
        GetHotZoneExpMultiplier = CSng(SvrConfig.GetValue("HotZoneExpMult"))
        If GetHotZoneExpMultiplier <= 0 Then GetHotZoneExpMultiplier = 1
    Else
        GetHotZoneExpMultiplier = 1
    End If
End Function
