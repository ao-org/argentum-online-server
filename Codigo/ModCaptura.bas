Attribute VB_Name = "ModCaptura"
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


Public Const CAPTURA_TIEMPO_ESPERA = 180 'Tiempo que dura la inscripcion

#If DEBUGGING Then
    Public Const CAPTURA_TIEMPO_INICIO_RONDA = 10 '60 'Tiempo hasta que se inicia la ronda
#Else
    Public Const CAPTURA_TIEMPO_INICIO_RONDA = 60 '60 'Tiempo hasta que se inicia la ronda
#End If
Public Const CAPTURA_TIEMPO_BANDERA = 10 'Tiempo que tiene que estar el user con la bandera en la base
Public Const CAPTURA_TIEMPO_MUERTE = 7 'Tiempo que tarda para poder revivir cuando muere
Public Const CAPTURA_TIEMPO_POR_MUERTE = 3 'Multiplicador de tiempo x veces que murio

Public Const MAP_SALA_ESPERA As Integer = 278
Public Const MAP_NEUTRAL As Integer = 276

Public Const MAP_TEAM_1 As Integer = 275
Public Const X_TEAM_1 As Integer = 43
Public Const Y_TEAM_1 As Integer = 51
Public Const X_BANDERA_1 As Integer = 37
Public Const Y_BANDERA_1 As Integer = 51

Public Const MAP_TEAM_2 As Integer = 277
Public Const X_TEAM_2 As Integer = 65
Public Const Y_TEAM_2 As Integer = 55
Public Const X_BANDERA_2 As Integer = 71
Public Const Y_BANDERA_2 As Integer = 55

Public Const MIN_SALA_ESPERA_X As Byte = 45
Public Const MIN_SALA_ESPERA_Y As Byte = 68
Public Const MAX_SALA_ESPERA_X As Byte = 56
Public Const MAX_SALA_ESPERA_Y As Byte = 73


Public Const OBJ_CAPTURA_BANDERA_1 As Integer = 3674 'Estandarte Azul
Public Const OBJ_CAPTURA_BANDERA_2 As Integer = 3675 'Estandarte Rojo


Public InstanciaCaptura As clsCaptura


