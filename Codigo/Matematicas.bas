Attribute VB_Name = "Matematicas"
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

Public Function Porcentaje(ByVal Total As Double, ByVal Porc As Double) As Double
        
        On Error GoTo Porcentaje_Err
        
100     Porcentaje = (Total * Porc) / 100

        
        Exit Function

Porcentaje_Err:
102     Call RegistrarError(Err.Number, Err.Description, "Matematicas.Porcentaje", Erl)
104
        
End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
        'Encuentra la distancia entre dos WorldPos
        
        On Error GoTo Distancia_Err
        
100     Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100&)

        
        Exit Function

Distancia_Err:
102     Call RegistrarError(Err.Number, Err.Description, "Matematicas.Distancia", Erl)
104
        
End Function

Function Distance(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant) As Double
        
        On Error GoTo Distance_Err
        

        'Encuentra la distancia entre dos puntos

100     Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

        
        Exit Function

Distance_Err:
102     Call RegistrarError(Err.Number, Err.Description, "Matematicas.Distance", Erl)
104
        
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero
        'Last Modify Date: 3/06/2006
        'Generates a random number in the range given - recoded to use longs and work properly with ranges
        '**************************************************************
        
        On Error GoTo RandomNumber_Err
        
100     RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

        
        Exit Function

RandomNumber_Err:
102     Call RegistrarError(Err.Number, Err.Description, "Matematicas.RandomNumber", Erl)
104
        
End Function
