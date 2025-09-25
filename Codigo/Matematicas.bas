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
    x As Double
    y As Double
End Type

Function max(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo max_Err
    If a > b Then
        max = a
    Else
        max = b
    End If
    Exit Function
max_Err:
    Call TraceError(Err.Number, Err.Description, "General.max", Erl)
End Function

Function Min(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo min_Err
    If a < b Then
        Min = a
    Else
        Min = b
    End If
    Exit Function
min_Err:
    Call TraceError(Err.Number, Err.Description, "General.min", Erl)
End Function

Public Function Porcentaje(ByVal total As Double, ByVal Porc As Double) As Double
    On Error GoTo Porcentaje_Err
    Porcentaje = (total * Porc) / 100
    Exit Function
Porcentaje_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Porcentaje", Erl)
End Function

Function Distancia(ByRef wp1 As t_WorldPos, ByRef wp2 As t_WorldPos) As Long
    On Error GoTo Distancia_Err
    Distancia = Abs(wp1.x - wp2.x) + Abs(wp1.y - wp2.y) + (Abs(wp1.Map - wp2.Map) * 100&)
    Exit Function
Distancia_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Distancia", Erl)
End Function

Function Distance(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
    On Error GoTo Distance_Err
    Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
    Exit Function
Distance_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.Distance", Erl)
End Function

Function GetDirection(ByRef From As t_WorldPos, ByRef ToPos As t_WorldPos) As t_Vector
    Dim Ret As t_Vector
    Ret.x = ToPos.x - From.x
    Ret.y = ToPos.y - From.y
    GetDirection = Ret
End Function

Function GetNormal(ByRef Vector As t_Vector) As t_Vector
    Dim length As Double
    Dim Ret    As t_Vector
    length = Distance(0, 0, Vector.x, Vector.y)
    Debug.Assert length <> 0
    Ret.x = Vector.x / length
    Ret.y = Vector.y / length
    GetNormal = Ret
End Function

Public Function ToRadians(ByVal degree As Double) As Double
    ToRadians = degree * PI / 180
End Function

Public Function RotateVector(ByRef v As t_Vector, ByVal angle As Double) As t_Vector
    Dim cosAngle As Double
    Dim sinAngle As Double
    Dim NewX     As Double
    Dim NewY     As Double
    cosAngle = Cos(angle)
    sinAngle = Sin(angle)
    NewX = v.x * cosAngle - v.y * sinAngle
    NewY = v.x * sinAngle + v.y * cosAngle
    RotateVector.x = NewX
    RotateVector.y = NewY
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    On Error GoTo RandomNumber_Err
    RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
    Exit Function
RandomNumber_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
End Function

Public Function RandomRange(ByVal LowerBound As Single, ByVal UpperBound As Single) As Single
    On Error GoTo RandomNumber_Err
    RandomRange = Rnd
    RandomRange = RandomRange * (UpperBound - LowerBound) + LowerBound
    Exit Function
RandomNumber_Err:
    Call TraceError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
End Function

Public Sub SetMask(ByRef Mask As Long, ByVal value As Long)
    Mask = Mask Or value
End Sub

Public Function IsSet(ByVal Mask As Long, ByVal value As Long) As Boolean
    IsSet = (Mask And value) <> 0
End Function

Public Sub UnsetMask(ByRef Mask As Long, ByVal value As Long)
    Mask = Mask And Not value
End Sub

Public Sub ResetMask(ByRef Mask As Long)
    Mask = 0
End Sub

Public Sub SetIntMask(ByRef Mask As Integer, ByVal value As Integer)
    Mask = Mask Or value
End Sub

Public Function IsIntSet(ByVal Mask As Integer, ByVal value As Integer) As Boolean
    IsIntSet = (Mask And value) <> 0
End Function

Public Sub UnsetIntMask(ByRef Mask As Integer, ByVal value As Integer)
    Mask = Mask And Not value
End Sub

Public Sub ResetIntMask(ByRef Mask As Integer)
    Mask = 0
End Sub

Public Function ShiftRight(ByVal Number As Long, ByVal BitCount As Byte) As Long
    If BitCount < 0 Or BitCount > 31 Then
        ShiftRight = 0
    Else
        ShiftRight = Number \ (2 ^ BitCount)
    End If
End Function

Public Function ShiftLeft(ByVal Number As Long, ByVal BitCount As Byte) As Long
    If BitCount < 0 Or BitCount > 31 Then
        ShiftLeft = 0
    ElseIf BitCount = 31 Then
        ' Directly assign the sign bit to avoid CLng overflow
        ShiftLeft = &H80000000
    Else
        ShiftLeft = Number * (2 ^ BitCount)
    End If
End Function
