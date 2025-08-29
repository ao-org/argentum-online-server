Attribute VB_Name = "Matematicas"
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

Const PI As Double = 3.14159265
Public Type t_Vector
    X As Double
    Y As Double
End Type

Function max(ByVal a As Double, ByVal b As Double) As Double
    On Error Goto max_Err
        On Error GoTo max_Err
100     If a > b Then
102         max = a
        Else
104         max = b
        End If
        Exit Function
max_Err:
106     Call TraceError(Err.Number, Err.Description, "General.max", Erl)
    Exit Function
max_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.max", Erl)
End Function
Function Min(ByVal a As Double, ByVal b As Double) As Double
    On Error Goto Min_Err
        On Error GoTo min_Err
100     If a < b Then
102         Min = a
        Else
104         Min = b
        End If
        Exit Function
min_Err:
106     Call TraceError(Err.Number, Err.Description, "General.min", Erl)
    Exit Function
Min_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Min", Erl)
End Function
Public Function Porcentaje(ByVal Total As Double, ByVal Porc As Double) As Double
    On Error Goto Porcentaje_Err
        On Error GoTo Porcentaje_Err
100     Porcentaje = (Total * Porc) / 100
        Exit Function
Porcentaje_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.Porcentaje", Erl)
    Exit Function
Porcentaje_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Porcentaje", Erl)
End Function
Function Distancia(ByRef wp1 As t_WorldPos, ByRef wp2 As t_WorldPos) As Long
    On Error Goto Distancia_Err
        On Error GoTo Distancia_Err
100     Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100&)
        Exit Function
Distancia_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.Distancia", Erl)
    Exit Function
Distancia_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Distancia", Erl)
End Function
Function Distance(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
    On Error Goto Distance_Err
        On Error GoTo Distance_Err
100     Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
        Exit Function
Distance_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.Distance", Erl)
    Exit Function
Distance_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Distance", Erl)
End Function

Function GetDirection(ByRef From As t_WorldPos, ByRef ToPos As t_WorldPos) As t_Vector
    On Error Goto GetDirection_Err
    Dim Ret As t_Vector
    Ret.x = ToPos.x - From.x
    Ret.y = ToPos.y - From.y
    GetDirection = Ret
    Exit Function
GetDirection_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.GetDirection", Erl)
End Function

Function GetNormal(ByRef Vector As t_Vector) As t_Vector
    On Error Goto GetNormal_Err
    Dim Length As Double
    Dim ret As t_Vector
    Length = distance(0, 0, Vector.X, Vector.Y)
    Debug.Assert length <> 0
    ret.X = Vector.X / Length
    ret.Y = Vector.Y / Length
    GetNormal = ret
    Exit Function
GetNormal_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.GetNormal", Erl)
End Function

Public Function ToRadians(ByVal degree As Double) As Double
    On Error Goto ToRadians_Err
    ToRadians = degree * PI / 180
    Exit Function
ToRadians_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.ToRadians", Erl)
End Function

Public Function RotateVector(ByRef v As t_Vector, ByVal angle As Double) As t_Vector
    On Error Goto RotateVector_Err
    Dim cosAngle As Double
    Dim sinAngle As Double
    Dim newX As Double
    Dim newY As Double
    
    cosAngle = Cos(angle)
    sinAngle = Sin(angle)
    
    newX = v.x * cosAngle - v.y * sinAngle
    newY = v.x * sinAngle + v.y * cosAngle
    
    RotateVector.x = newX
    RotateVector.y = newY
    Exit Function
RotateVector_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.RotateVector", Erl)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    On Error Goto RandomNumber_Err
        On Error GoTo RandomNumber_Err
100     RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
        Exit Function
RandomNumber_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
    Exit Function
RandomNumber_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
End Function

Public Function RandomRange(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
    On Error Goto RandomRange_Err
        On Error GoTo RandomNumber_Err
100     RandomRange = Rnd
        RandomRange = RandomRange * (UpperBound - LowerBound) + LowerBound
        Exit Function
RandomNumber_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
    Exit Function
RandomRange_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.RandomRange", Erl)
End Function

Public Sub SetMask(ByRef mask As Long, ByVal value As Long)
    On Error Goto SetMask_Err
    mask = mask Or value
    Exit Sub
SetMask_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.SetMask", Erl)
End Sub

Public Function IsSet(ByVal Mask As Long, ByVal Value As Long) As Boolean
    On Error Goto IsSet_Err
    IsSet = (Mask And value) <> 0
    Exit Function
IsSet_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.IsSet", Erl)
End Function

Public Sub UnsetMask(ByRef mask As Long, ByVal value As Long)
    On Error Goto UnsetMask_Err
    mask = mask And Not value
    Exit Sub
UnsetMask_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.UnsetMask", Erl)
End Sub

Public Sub ResetMask(ByRef mask As Long)
    On Error Goto ResetMask_Err
    mask = 0
    Exit Sub
ResetMask_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.ResetMask", Erl)
End Sub

Public Sub SetIntMask(ByRef Mask As Integer, ByVal Value As Integer)
    On Error Goto SetIntMask_Err
    Mask = Mask Or Value
    Exit Sub
SetIntMask_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.SetIntMask", Erl)
End Sub

Public Function IsIntSet(ByVal Mask As Integer, ByVal value As Integer) As Boolean
    On Error Goto IsIntSet_Err
    IsIntSet = (Mask And value) <> 0
    Exit Function
IsIntSet_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.IsIntSet", Erl)
End Function

Public Sub UnsetIntMask(ByRef Mask As Integer, ByVal Value As Integer)
    On Error Goto UnsetIntMask_Err
    Mask = Mask And Not Value
    Exit Sub
UnsetIntMask_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.UnsetIntMask", Erl)
End Sub

Public Sub ResetIntMask(ByRef Mask As Integer)
    On Error Goto ResetIntMask_Err
    Mask = 0
    Exit Sub
ResetIntMask_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.ResetIntMask", Erl)
End Sub

Public Function ShiftRight(ByVal Number As Long, ByVal BitCount As Byte) As Long
    On Error Goto ShiftRight_Err
    If BitCount < 0 Or BitCount > 31 Then
        ShiftRight = 0
    Else
        ShiftRight = Number \ (2 ^ BitCount)
    End If
    Exit Function
ShiftRight_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.ShiftRight", Erl)
End Function

Public Function ShiftLeft(ByVal Number As Long, ByVal BitCount As Byte) As Long
    On Error Goto ShiftLeft_Err
    If BitCount < 0 Or BitCount > 31 Then
        ShiftLeft = 0
    ElseIf BitCount = 31 Then
        ' Directly assign the sign bit to avoid CLng overflow
        ShiftLeft = &H80000000
    Else
        ShiftLeft = Number * (2 ^ BitCount)
    End If
    Exit Function
ShiftLeft_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.ShiftLeft", Erl)
End Function
