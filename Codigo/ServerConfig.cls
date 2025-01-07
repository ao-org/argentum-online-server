VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServerConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server
'
'    Copyright (C) 2025 Noland Studios LTD
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
Option Base 0

Private mSettings As Dictionary

Public Function GetValue(ByVal key As String) As Variant
    Debug.Assert Not mSettings Is Nothing
    Debug.Assert mSettings.Exists(key)
    If mSettings.Exists(key) Then
        GetValue = mSettings.Item(key)
    End If
End Function

Private Sub DumpSettings()
    Debug.Assert Not mSettings Is Nothing
    Dim key
    Dim value
    For Each key In mSettings.Keys
        value = mSettings.Item(key)
        Debug.Print "[Settings] " & key & " -> " & value
    Next key
End Sub

Public Function LoadSettings(ByVal Filename As String) As Long
    LoadSettings = -1
    Dim reader As clsIniManager
    Set reader = New clsIniManager
    Call reader.Initialize(Filename)
    
    mSettings.Add "ExpMult", CInt(val(reader.GetValue("CONFIGURACIONES", "ExpMult")))
    mSettings.Add "GoldMult", CInt(val(reader.GetValue("CONFIGURACIONES", "OroMult")))
    mSettings.Add "RecoleccionMult", CDbl(val(reader.GetValue("CONFIGURACIONES", "RecoleccionMult")))
    mSettings.Add "OroPorNivelBilletera", CLng(val(reader.GetValue("CONFIGURACIONES", "OroPorNivelBilletera")))
    mSettings.Add "OroPorNivel", val(reader.GetValue("CONFIGURACIONES", "OroPorNivel"))
    mSettings.Add "DayLength", CLng(val(reader.GetValue("CONFIGURACIONES", "DuracionDia"))) * 60 * 1000
    mSettings.Add "CostoPerdonPorCiudadano", CLng(val(reader.GetValue("CONFIGURACIONES", "CostoPerdonPorCiudadano")))
    mSettings.Add "ChatGlobal", CInt(val(reader.GetValue("CONFIGURACIONES", "ChatGlobal", "1")))
    
    mSettings.Add "DropMult", CInt(val(reader.GetValue("DROPEO", "DropMult")))
    mSettings.Add "DropActive", CInt(val(reader.GetValue("DROPEO", "DropActive")))
    
    mSettings.Add "MaximoSpeedHack", CInt(val(reader.GetValue("ANTICHEAT", "MaximoSpeedHack")))
    
    mSettings.Add "FISHING_REQUIRED_PERCENT", CInt(val(reader.GetValue("PESCA", "MinAgua")))
    mSettings.Add "FISHING_TILES_ON_MAP", CInt(val(reader.GetValue("PESCA", "MaxZonas")))
    mSettings.Add "FISHING_POOL_ID", CInt(val(reader.GetValue("PESCA", "FishingID")))

    Set reader = Nothing
    
    LoadSettings = 0
    DumpSettings
End Function

Private Sub Class_Initialize()
    Set mSettings = New Dictionary
End Sub
