Attribute VB_Name = "Consts"
' Argentum 20 Game Server
'
'    Copyright (C) 2025-2026 Noland Studios LTD
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

' AI Default values
Public Const DEFAULT_NPC_VISION_RANGE_X             As Integer = 15
Public Const DEFAULT_NPC_VISION_RANGE_Y             As Integer = 13
Public Const DEFAULT_NPC_SPELL_RANGE_X              As Integer = 11
Public Const DEFAULT_NPC_SPELL_RANGE_Y              As Integer = 9
Public Const DEFAULT_NPC_MAX_VISION_RANGE           As Integer = 100
Public Const DEFAULT_ORBIT_REEVALUATE_MS            As Long = 1800
Public Const DEFAULT_PATH_RECOMPUTE_COOLDOWN_MS     As Long = 250
Public Const DEFAULT_NPC_ORBIT_TANGENT_WEIGHT       As Double = 0.35
Public Const DEFAULT_NPC_RETREAT_DISTANCE_BUFFER    As Double = 0.75
Public Const DEFAULT_NPC_ORBIT_STEP_DEGREES         As Double = 55
Public Const DEFAULT_NPC_STRAFE_DURATION_MS         As Long = 900
Public Const DEFAULT_NPC_WANDER_BLOCKED_ATTEMPTS    As Integer = 2
Public Const DEFAULT_NPC_WANDER_DEFAULT_MIN_STEPS   As Integer = 3
Public Const DEFAULT_NPC_WANDER_DEFAULT_MAX_STEPS   As Integer = 6
Public Const DEFAULT_NPC_WANDER_DEFAULT_MIN_PAUSE   As Long = 1000
Public Const DEFAULT_NPC_WANDER_DEFAULT_MAX_PAUSE   As Long = 3000
Public Const DEFAULT_NPC_WANDER_DEFAULT_RADIUS      As Integer = 6
Public Const DEFAULT_NPC_WANDER_ANIMAL_MIN_STEPS    As Integer = 4
Public Const DEFAULT_NPC_WANDER_ANIMAL_MAX_STEPS    As Integer = 8
Public Const DEFAULT_NPC_WANDER_ANIMAL_MIN_PAUSE    As Long = 500
Public Const DEFAULT_NPC_WANDER_ANIMAL_MAX_PAUSE    As Long = 1800
Public Const DEFAULT_NPC_WANDER_ANIMAL_RADIUS       As Integer = 8
Public Const DEFAULT_NPC_WANDER_GUARD_MIN_STEPS     As Integer = 2
Public Const DEFAULT_NPC_WANDER_GUARD_MAX_STEPS     As Integer = 4
Public Const DEFAULT_NPC_WANDER_GUARD_MIN_PAUSE     As Long = 2000
Public Const DEFAULT_NPC_WANDER_GUARD_MAX_PAUSE     As Long = 5000
Public Const DEFAULT_NPC_WANDER_GUARD_RADIUS        As Integer = 4



'General consts
Public Const MAX_INTEGER                            As Integer = 32767
Public Const MAX_LONG                               As Long = 2147483647


'FX
Public Const FX_STABBING                            As Byte = 89
Public Const FX_BLOOD                               As Byte = 14


'TOWN NAMES
Public Const CIUDAD_ULLATHORPE                      As String = "Ullathorpe"
Public Const CIUDAD_NIX                             As String = "Nix"
Public Const CIUDAD_BANDERBILL                      As String = "Banderbill"
Public Const CIUDAD_LINDOS                          As String = "Lindos"
Public Const CIUDAD_ARGHAL                          As String = "Arghal"
Public Const CIUDAD_FORGAT                          As String = "Forgat"
Public Const CIUDAD_ARKHEIN                         As String = "Arkhein"
Public Const CIUDAD_ELDORIA                         As String = "Eldoria"
Public Const CIUDAD_PENTHAR                         As String = "Penthar"
Public Const CIUDAD_MORGRIM                         As String = "Morgrim"
