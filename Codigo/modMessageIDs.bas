Attribute VB_Name = "modMessageIDs"
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
    
    
Public Const MSG_PARALYZED                                  As Integer = 1123
Public Const MSG_ALREADY_HIDDEN                             As Integer = 1127
Public Const MSG_CANNOT_HIDE_MOUNTED                        As Integer = 704
Public Const MSG_HIDE_FAILED                                As Integer = 57
Public Const MSG_TAME_FAILED                                As Integer = 659
Public Const MSG_MAP_NEWBIE_ONLY                            As Integer = 771
Public Const MSG_MAP_ONLY_CITIZENS                          As Integer = 772
Public Const MSG_MAP_ONLY_CRIMINALS                         As Integer = 773
Public Const MSG_MAP_REQUIRES_CLAN                          As Integer = 774
Public Const MSG_MAP_MIN_LEVEL                              As Integer = 1108
Public Const MSG_MAP_MAX_LEVEL                              As Integer = 1109
Public Const MSG_MAP_REQUIRES_GROUP                         As Integer = 775
Public Const MSG_MAP_REQUIRES_PATREON                       As Integer = 776
Public Const MSG_TILE_REQUIRES_PATREON                      As Integer = 776
Public Const MSG_PICKUP_UNAVAILABLE                         As Integer = 261

Public Const MSG_CLIENT_UPDATE_REQUIRED                        As Integer = 2092
Public Const MSG_INVALID_SESSION_TOKEN                         As Integer = 2093
Public Const MSG_CONNECTION_SLOT_ERROR                         As Integer = 2094
Public Const MSG_DISABLED_NEW_CHARACTERS                       As Integer = 1776
Public Const MSG_YOU_HAVE_TOO_MANY_CHARS                       As Integer = 1777
Public Const MSG_USERNAME_ALREADY_TAKEN                        As Integer = 1778
Public Const MSG_UPGRADE_ACCOUNT_TO_CREATE_MORE_CHARS          As Integer = 1779
Public Const MSG_RIDER_LEVEL_REQUIREMENT                       As Integer = 2078
Public Const MSG_CHARACTER_HOME                                As Integer = 2104
Public Const MSG_QUEST_LEVEL_REQUIREMENT                       As Integer = 1426
Public Const MSG_NPC_INMUNE_TO_SPELLS                          As Integer = 666
Public Const MSG_BLODIUM_ANVIL_REQUIRED                        As Integer = 2113
Public Const MSG_CANNOT_PICK_UP_ITEMS_IN_JAIL                  As Integer = 2109
Public Const MSG_CANNOT_DROP_ITEMS_IN_JAIL                     As Integer = 2110
Public Const MSG_CANNOT_TRADE_IN_JAIL                          As Integer = 2111
Public Const MSG_CANNOT_USE_HOME_IN_JAIL                       As Integer = 2112
Public Const MAP_HOME_IN_JAIL                                  As Integer = 66
Public Const MSG_QUEST_ALREADY_COMPLETED                       As Integer = 2114

' Msg2117 = You need to have ¬1 equipped to perform this action.
Public Const MSG_REMOVE_NEED_EQUIPPED As Integer = 2117

' Msg2118 = Pay attention! You lost your net, it got caught on the special fish.
Public Const MSG_REMOVE_NET_LOST As Integer = 2118

' Msg2119 = Pay attention! You almost lost your net to the special fish.
Public Const MSG_REMOVE_NET_ALMOST_LOST As Integer = 2119

' Msg2120 = Pay attention! You almost lost your fishing rod to the special fish.
Public Const MSG_REMOVE_ALMOST_YOUR_FISHING As Integer = 2120

' Message IDs used for faction connection notifications (randomized variants).
Public Const MSG_CONNECTION_ROYAL_ARMY_1 As Integer = 2133
Public Const MSG_CONNECTION_ROYAL_ARMY_2 As Integer = 2134
Public Const MSG_CONNECTION_ROYAL_ARMY_3 As Integer = 2135
Public Const MSG_CONNECTION_ROYAL_ARMY_4 As Integer = 2136
Public Const MSG_CONNECTION_ROYAL_ARMY_5 As Integer = 2137
Public Const MSG_CONNECTION_ROYAL_ARMY_6 As Integer = 2138
Public Const MSG_CONNECTION_ROYAL_ARMY_7 As Integer = 2139
Public Const MSG_CONNECTION_ROYAL_ARMY_8 As Integer = 2140
Public Const MSG_CONNECTION_ROYAL_ARMY_9 As Integer = 2141
Public Const MSG_CONNECTION_ROYAL_ARMY_10 As Integer = 2142
Public Const MSG_CONNECTION_DARK_LEGION_1 As Integer = 2149
Public Const MSG_CONNECTION_DARK_LEGION_2 As Integer = 2150
Public Const MSG_CONNECTION_DARK_LEGION_3 As Integer = 2151
Public Const MSG_CONNECTION_DARK_LEGION_4 As Integer = 2152
Public Const MSG_CONNECTION_DARK_LEGION_5 As Integer = 2153
Public Const MSG_CONNECTION_DARK_LEGION_6 As Integer = 2154
Public Const MSG_CONNECTION_DARK_LEGION_7 As Integer = 2155
Public Const MSG_CONNECTION_DARK_LEGION_8 As Integer = 2156
Public Const MSG_CONNECTION_DARK_LEGION_9 As Integer = 2157
Public Const MSG_CONNECTION_DARK_LEGION_10 As Integer = 2158
Public Const MSG_PERFORATED_ARMOR As Integer = 2161
