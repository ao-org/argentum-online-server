Attribute VB_Name = "ModClimas"
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
Public IntensidadDeNubes   As Byte

Public IntensidadDeLluvias As Byte

Public CapasLlueveEn       As Integer

Public TimerMeteorologico  As Byte

Public DuracionDeLLuvia    As Integer

Public ServidorNublado     As Boolean

Public ProbabilidadNublar  As Byte

Public ProbabilidadLLuvia  As Byte

Public Sub ResetMeteo()
        
        On Error GoTo ResetMeteo_Err
        
100     Call AgregarAConsola("Servidor > Meteorologia reseteada")
102     frmMain.TimerMeteorologia.Enabled = True
104     frmMain.Truenos.Enabled = False
106     TimerMeteorologico = 30
108     ServidorNublado = False
110     Lloviendo = False

        
        Exit Sub

ResetMeteo_Err:
112     Call TraceError(Err.Number, Err.Description, "ModClimas.ResetMeteo", Erl)

        
End Sub
