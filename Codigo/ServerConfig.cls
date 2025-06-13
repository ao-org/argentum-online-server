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
    
    'The number of outstanding receive buffers allocated for each DirectPlay thread. If you increase the number of receive buffers, DirectPlay can pull more data out of the operating system buffers. However, you may also increase latency if data is arriving faster than your application can process it.
    mSettings.Add "DP_BuffersPerThread", CLng(val(reader.GetValue("DIRECTPLAY", "BuffersPerThread", 5)))
    
    mSettings.Add "DP_DefaultEnumCount", CLng(val(reader.GetValue("DIRECTPLAY", "DefaultEnumCount", 5)))
    'Long value that specifies the default retry interval, in milliseconds.
    mSettings.Add "DP_DefaultEnumRetryInterval", CLng(val(reader.GetValue("DIRECTPLAY", "DefaultEnumRetryInterval", 100)))
    mSettings.Add "DP_DefaultEnumTimeout", CLng(val(reader.GetValue("DIRECTPLAY", "DefaultEnumTimeout", 500)))
    mSettings.Add "DP_MaxEnumPayloadSize", CLng(val(reader.GetValue("DIRECTPLAY", "MaxEnumPayloadSize", 1024)))

    'Number of threads the service provider will use for servicing network requests. The default value is based on an algorithm that takes into account the number of processors on the system.
    'Most applications will not need to modify this value. After a service provider is active in your process you may only increase this value.
    'Decreasing the value will have no effect. The setting is process wide, meaning it will effect your current Microsoft® DirectPlay® object and any other DirectPlay objects in your process.
    mSettings.Add "DP_NumThreads", CLng(val(reader.GetValue("DIRECTPLAY", "NumThreads", 48)))
    mSettings.Add "DP_SystemBufferSizeAs", CLng(val(reader.GetValue("DIRECTPLAY", "SystemBufferSizeAs", 65536)))
    
    'Number of milliseconds DirectPlay should wait, since the last time it received a packet from an endpoint, before it sends a keep-alive message. You can catch the KeepAlive up to 2 times of what you specify.
    mSettings.Add "DP_TimeoutUntilKeepAlive", CLng(val(reader.GetValue("DIRECTPLAY", "TimeoutUntilKeepAlive", 500)))
    
    'Number of connection retries DirectPlay should make during the connection process.
    mSettings.Add "DP_ConnectRetries", CLng(val(reader.GetValue("DIRECTPLAY", "ConnectRetries", 20)))
    
    'Number of milliseconds DirectPlay should wait before it retries a connection request.
    mSettings.Add "DP_ConnectTimeout", CLng(val(reader.GetValue("DIRECTPLAY", "ConnectTimeout", 500)))


    

    
    mSettings.Add "ExpMult", CInt(val(reader.GetValue("CONFIGURACIONES", "ExpMult")))
    mSettings.Add "GoldMult", CInt(val(reader.GetValue("CONFIGURACIONES", "OroMult")))
    mSettings.Add "RecoleccionMult", CDbl(val(reader.GetValue("CONFIGURACIONES", "RecoleccionMult")))
    mSettings.Add "OroPorNivelBilletera", CLng(val(reader.GetValue("CONFIGURACIONES", "OroPorNivelBilletera")))
    mSettings.Add "OroPorNivel", val(reader.GetValue("CONFIGURACIONES", "OroPorNivel"))
    mSettings.Add "DayLength", CLng(val(reader.GetValue("CONFIGURACIONES", "DuracionDia"))) * 60 * 1000
    mSettings.Add "CostoPerdonPorCiudadano", CLng(val(reader.GetValue("CONFIGURACIONES", "CostoPerdonPorCiudadano")))
    mSettings.Add "ChatGlobal", CInt(val(reader.GetValue("CONFIGURACIONES", "ChatGlobal", "1")))
    mSettings.Add "PartyELV", CInt(val(Reader.GetValue("CONFIGURACIONES", "PartyELV", "4")))
    mSettings.Add "PenaltyExpUserPerLevel", CSng(val(reader.GetValue("CONFIGURACIONES", "PenaltyExpUserPerLevel", "0.05")))
    mSettings.Add "DeltaLevelExpPenalty", CInt(val(Reader.GetValue("CONFIGURACIONES", "DeltaLevelExpPenalty", "4")))
    mSettings.Add "LeadershipExpPartyBonus", CDbl(val(reader.GetValue("CONFIGURACIONES", "LeadershipExpPartyBonus", "1.01")))

    mSettings.Add "MaxJailTime", CInt(val(reader.GetValue("CONFIGURACIONES", "MaxJailTime", "8640")))
    
    mSettings.Add "DropMult", CInt(val(reader.GetValue("DROPEO", "DropMult")))
    mSettings.Add "DropActive", CInt(val(reader.GetValue("DROPEO", "DropActive")))
    
    mSettings.Add "MaximoSpeedHack", CInt(val(reader.GetValue("ANTICHEAT", "MaximoSpeedHack")))
    
    mSettings.Add "FISHING_REQUIRED_PERCENT", CInt(val(reader.GetValue("PESCA", "MinAgua")))
    mSettings.Add "FISHING_TILES_ON_MAP", CInt(val(reader.GetValue("PESCA", "MaxZonas")))
    mSettings.Add "FISHING_POOL_ID", CInt(val(reader.GetValue("PESCA", "FishingID")))
    mSettings.Add "FISHING_MAP_SPECIAL_FISH1_ID", CInt(val(reader.GetValue("PESCA", "FishingMapSpecialFishID")))
    mSettings.Add "FISHING_SPECIALFISH1_ID", CInt(val(reader.GetValue("PESCA", "UniqueMapfish1")))
    mSettings.Add "FISHING_SPECIALFISH2_ID", CInt(val(reader.GetValue("PESCA", "UniqueMapfish2")))
    mSettings.Add "FISHING_SPECIALFISH1_REMPLAZO_ID", CInt(val(reader.GetValue("PESCA", "ReplacementSpecialFish")))
    mSettings.Add "NPC_WATCHMAN_ID", CInt(val(reader.GetValue("PESCA", "NPCFishingWatcher")))
    

    mSettings.Add "NPC_SPELL_RANGE_X", CInt(val(Reader.GetValue("AI", "NPCSpellRangeX", DEFAULT_NPC_SPELL_RANGE_X)))
    mSettings.Add "NPC_SPELL_RANGE_Y", CInt(val(Reader.GetValue("AI", "NPCSpellRangeY", DEFAULT_NPC_SPELL_RANGE_Y)))
    mSettings.Add "NPC_VISION_RANGE_X", CInt(val(Reader.GetValue("AI", "NPCVisionRangeX", DEFAULT_NPC_VISION_RANGE_X)))
    mSettings.Add "NPC_VISION_RANGE_Y", CInt(val(Reader.GetValue("AI", "NPCVisionRangeY", DEFAULT_NPC_VISION_RANGE_Y)))
    mSettings.Add "NPC_PATHFINDING_MAX_STEPS", Min(MAX_PATH_LENGTH, CInt(val(Reader.GetValue("AI", "NPCMaxStepsPathFinding", MAX_PATH_LENGTH))))
    mSettings.Add "NPC_MAX_VISION_RANGE", Min(DEFAULT_NPC_MAX_VISION_RANGE, CInt(val(Reader.GetValue("AI", "NPCMaxVisionRange", DEFAULT_NPC_MAX_VISION_RANGE))))
      
    Set reader = Nothing
    
    LoadSettings = 0
    DumpSettings
End Function

Private Sub Class_Initialize()
    Set mSettings = New Dictionary
End Sub
