Attribute VB_Name = "ModClimas"
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
'
Public IntensidadDeNubes   As Byte
Public IntensidadDeLluvias As Byte
Public CapasLlueveEn       As Integer
Public TimerMeteorologico  As Byte
Public DuracionDeLLuvia    As Integer
Public ServidorNublado     As Boolean
Public ProbabilidadNublar  As Byte
Public ProbabilidadLLuvia  As Byte

Public Function IsAtmosphericFogActive() As Boolean
    IsAtmosphericFogActive = ServidorNublado Or Nieblando
End Function

Public Sub DetermineInitialWeatherSynchronization(ByRef SendRain As Boolean, ByRef SendSnow As Boolean, ByRef ToggleFog As Boolean)
    SendRain = Lloviendo
    SendSnow = Nebando
    ToggleFog = IsAtmosphericFogActive()
End Sub

Public Sub DetermineWeatherResetNotifications(ByVal WasRaining As Boolean, ByVal WasSnowing As Boolean, ByVal WasFogActive As Boolean, ByRef NotifyRain As Boolean, ByRef NotifySnow As Boolean, ByRef NotifyFog As Boolean)
    NotifyRain = WasRaining
    NotifySnow = WasSnowing
    NotifyFog = WasFogActive
End Sub

Public Sub ResetMeteo(Optional ByVal NotifyClients As Boolean = False)
    On Error GoTo ResetMeteo_Err
    Dim NotifyRain As Boolean
    Dim NotifySnow As Boolean
    Dim NotifyFog As Boolean
    Call DetermineWeatherResetNotifications(Lloviendo, Nebando, IsAtmosphericFogActive(), NotifyRain, NotifySnow, NotifyFog)
    Call AgregarAConsola("Servidor > Meteorologia reseteada")
    frmMain.TimerMeteorologia.Enabled = True
    frmMain.Truenos.Enabled = False
    TimerMeteorologico = 30
    ServidorNublado = False
    Nieblando = False
    Lloviendo = False
    Nebando = False
    If NotifyClients Then
        ' Fog has no explicit state in its packet: only toggle clients that
        ' were known to have atmospheric fog active before this reset.
        If NotifyFog Then Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle(IntensidadDeNubes))
        If NotifyRain Then Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        If NotifySnow Then Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
    End If
    Exit Sub
ResetMeteo_Err:
    Call TraceError(Err.Number, Err.Description, "ModClimas.ResetMeteo", Erl)
End Sub
