Attribute VB_Name = "modFontManager"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
        .r = 0: .g = 130: .b = 255
    End With
    With FontTypeColors(e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA)
        .r = 0: .g = 100: .b = 255
    End With
    With FontTypeColors(e_FontTypeNames.FONTTYPE_CONSEJO)
        .r = 120: .g = 220: .b = 255
    End With
    With FontTypeColors(e_FontTypeNames.FONTTYPE_CRIMINAL)
        .r = 240: .g = 75: .b = 70
    End With
    With FontTypeColors(e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS)
        .r = 255: .g = 0: .b = 0
    End With
    With FontTypeColors(e_FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .r = 160: .g = 15: .b = 0
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
