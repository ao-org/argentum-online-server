Attribute VB_Name = "Matematicas"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Option Explicit
Function max(ByVal a As Double, ByVal b As Double) As Double
        On Error GoTo max_Err
100     If a > b Then
102         max = a
        Else
104         max = b
        End If
        Exit Function
max_Err:
106     Call TraceError(Err.Number, Err.Description, "General.max", Erl)
End Function
Function Min(ByVal a As Double, ByVal b As Double) As Double
        On Error GoTo min_Err
100     If a < b Then
102         Min = a
        Else
104         Min = b
        End If
        Exit Function
min_Err:
106     Call TraceError(Err.Number, Err.Description, "General.min", Erl)
End Function
Public Function Porcentaje(ByVal Total As Double, ByVal Porc As Double) As Double
        On Error GoTo Porcentaje_Err
100     Porcentaje = (Total * Porc) / 100
        Exit Function
Porcentaje_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.Porcentaje", Erl)
End Function
Function Distancia(ByRef wp1 As t_WorldPos, ByRef wp2 As t_WorldPos) As Long
        On Error GoTo Distancia_Err
100     Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100&)
        Exit Function
Distancia_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.Distancia", Erl)
End Function
Function Distance(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
        On Error GoTo Distance_Err
100     Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))
        Exit Function
Distance_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.Distance", Erl)
End Function
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        On Error GoTo RandomNumber_Err
100     RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound
        Exit Function
RandomNumber_Err:
102     Call TraceError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
End Function

Public Sub SetMask(ByRef mask As Long, ByVal value As Long)
    mask = mask Or value
End Sub

Public Function IsSet(ByVal mask As Long, ByVal value As Long)
    IsSet = (mask And value) > 0
End Function

Public Sub UnsetMask(ByRef mask As Long, ByVal value As Long)
    mask = mask And Not value
End Sub

Public Sub ResetMask(ByRef mask As Long)
    mask = 0
End Sub
