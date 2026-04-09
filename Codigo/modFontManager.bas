Attribute VB_Name = "modFontManager"
Type t_FontTypeColor
    r As Byte
    g As Byte
    b As Byte
End Type
Public Const MAX_FONTTYPES As Byte = 100
Public FontTypeColors(0 To MAX_FONTTYPES) As t_FontTypeColor
Public Function FontTypeToColor(ByVal fontType As e_FontTypeNames) As Long
    On Error GoTo FontTypeToColor_Err
    If fontType < 0 Or fontType > MAX_FONTTYPES Then
        FontTypeToColor = RGB(255, 255, 255)
        Exit Function
    End If

    With FontTypeColors(fontType)
        FontTypeToColor = RGB(.r, .g, .b)
    End With
    Exit Function
FontTypeToColor_Err:
    Call TraceError(Err.Number, Err.Description, "modFontManager.FontTypeToColor", Erl)
End Function
Public Sub InitFontTypeColors()
    On Error GoTo InitFontTypeColors_Err
    With FontTypeColors(e_FontTypeNames.FONTTYPE_CITIZEN)
        .r = 6: .g = 128: .b = 255
    End With

    With FontTypeColors(e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA)
        .r = 60: .g = 163: .b = 255
    End With

    With FontTypeColors(e_FontTypeNames.FONTTYPE_CRIMINAL)
        .r = 255: .g = 0: .b = 0
    End With

    With FontTypeColors(e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS)
        .r = 255: .g = 51: .b = 51
    End With

    With FontTypeColors(e_FontTypeNames.FONTTYPE_CONSEJO)
        .r = 66: .g = 201: .b = 255
    End With

    With FontTypeColors(e_FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .r = 255: .g = 102: .b = 102
    End With
    Exit Sub
InitFontTypeColors_Err:
    Call TraceError(Err.Number, Err.Description, "modFontManager.InitFontTypeColors", Erl)
End Sub
Public Function GetFontTypeByFactionStatus(ByVal Status As e_Facciones) As e_FontTypeNames
    On Error GoTo GetFontTypeByFactionStatus_Err
    Select Case Status
        Case e_Facciones.Criminal
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CRIMINAL

        Case e_Facciones.Caos
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS

        Case e_Facciones.concilio
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CONSEJOCAOS

        Case e_Facciones.Ciudadano
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CITIZEN

        Case e_Facciones.Armada
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA

        Case e_Facciones.consejo
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CONSEJO
        Case Else
            GetFontTypeByFactionStatus = e_FontTypeNames.FONTTYPE_CITIZEN
    End Select
    Exit Function
GetFontTypeByFactionStatus_Err:
    Call TraceError(Err.Number, Err.Description, "modFontManager.GetFontTypeByFactionStatus", Erl)
End Function
