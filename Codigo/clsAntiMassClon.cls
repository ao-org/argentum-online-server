VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAntiMassClon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const MaximoPersonajesPorIP = 15

Private m_coleccion As New Collection

Public Function MaxPersonajes(sIp As String) As Boolean
    On Error Goto MaxPersonajes_Err
        
        On Error GoTo MaxPersonajes_Err
        

        Dim i As Long

100     For i = 1 To m_coleccion.Count

102         If m_coleccion.Item(i).IP = sIp Then
104             m_coleccion.Item(i).PersonajesCreados = m_coleccion.Item(i).PersonajesCreados + 1
106             MaxPersonajes = (m_coleccion.Item(i).PersonajesCreados > MaximoPersonajesPorIP)

108             If MaxPersonajes Then m_coleccion.Item(i).PersonajesCreados = 16
                Exit Function

            End If

110     Next i

112     MaxPersonajes = False
        Exit Function

        
        Exit Function

MaxPersonajes_Err:
114     Call TraceError(Err.Number, Err.Description, "clsAntiMassClon.MaxPersonajes", Erl)
116
        
    Exit Function
MaxPersonajes_Err:
    Call TraceError(Err.Number, Err.Description, "clsAntiMassClon.MaxPersonajes", Erl)
End Function

Public Function VaciarColeccion()
    On Error Goto VaciarColeccion_Err

        On Error GoTo ErrHandler

        Dim i As Integer

100     For i = 1 To m_coleccion.Count
102         Call m_coleccion.Remove(1)
        Next

        Exit Function
ErrHandler:
104     Call LogError("Error en RestarConexion " & Err.Description)

    Exit Function
VaciarColeccion_Err:
    Call TraceError(Err.Number, Err.Description, "clsAntiMassClon.VaciarColeccion", Erl)
End Function

