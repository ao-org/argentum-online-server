Attribute VB_Name = "modTime"
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
'
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As t_SYSTEMTIME)

Private theTime      As t_SYSTEMTIME


Private Type t_SYSTEMTIME

    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Public Type t_Timer
    ElapsedTime As Long
    Interval As Long
    Occurrences As Integer
End Type

Public Function GetTickCount() As Long
        On Error GoTo GetTickCount_Err
        GetTickCount = timeGetTime And &H7FFFFFFF
        Exit Function
GetTickCount_Err:
        Call TraceError(Err.Number, Err.Description, "ModLadder.GetTickCount", Erl)
End Function

Function GetTimeFormated() As String
        On Error GoTo GetTimeFormated_Err
        Dim Elapsed As Long
        Elapsed = (GetTickCount() - HoraMundo) / DuracionDia
        Dim Mins As Long
        Mins = (Elapsed - Fix(Elapsed)) * 1440
        Dim Horita    As Byte
        Dim Minutitos As Byte
        Horita = Fix(Mins / 60)
        Minutitos = Mins Mod 60
        GetTimeFormated = Right$("00" & Horita, 2) & ":" & Right$("00" & Minutitos, 2)
        Exit Function
GetTimeFormated_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.GetTimeFormated - " + Erl, Erl)
End Function

Public Sub GetHoraActual()
        On Error GoTo GetHoraActual_Err
        GetSystemTime theTime
        HoraActual = (theTime.wHour - 3)
        If HoraActual = -3 Then HoraActual = 21
        If HoraActual = -2 Then HoraActual = 22
        If HoraActual = -1 Then HoraActual = 23
        frmMain.lblhora.Caption = HoraActual & ":" & Format(theTime.wMinute, "00") & ":" & Format(theTime.wSecond, "00")
        HoraEvento = HoraActual
        Exit Sub
GetHoraActual_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.GetHoraActual", Erl)
End Sub

Public Function SumarTiempo(segundos As Integer) As String
        On Error GoTo SumarTiempo_Err
        Dim a As Variant, b As Variant
        Dim X As Integer
        Dim T As String
        T = "00:00:00" 'Lo inicializamos en 0 horas, 0 minutos, 0 segundos
        a = Format("00:00:01", "hh:mm:ss") 'guardamos en una variable el formato de 1 segundos
        If segundos > 0 Then
             For X = 1 To segundos 'hacemos segundo a segundo
                b = Format(T, "hh:mm:ss") 'En B guardamos un formato de hora:minuto:segundo segun lo que tenia T
                T = Format(TimeValue(a) + TimeValue(b), "hh:mm:ss") 'asignamos a T la suma de A + B (osea, sumamos logicamente 1 segundo)
             Next X
        End If
        SumarTiempo = T 'a la funcion le damos el valor que hallamos en T
        Exit Function
SumarTiempo_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.SumarTiempo", Erl)
End Function

Public Sub SetTimer(ByRef timer As t_Timer, ByVal Interval As Long)
    timer.ElapsedTime = 0
    timer.Interval = Interval
    timer.Occurrences = 0
End Sub

Public Function UpdateTime(ByRef timer As t_Timer, ByVal deltaTime As Long) As Boolean
    timer.ElapsedTime = timer.ElapsedTime + deltaTime
    UpdateTime = timer.ElapsedTime - timer.Interval > 0
    timer.ElapsedTime = timer.ElapsedTime Mod timer.Interval
    If UpdateTime Then
        timer.Occurrences = timer.Occurrences + 1
    End If
End Function
