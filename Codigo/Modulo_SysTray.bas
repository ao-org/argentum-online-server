Attribute VB_Name = "SysTray"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                       SysTray
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Para minimizar a la barra de tareas
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Type CWPSTRUCT

    lParam As Long
    wParam As Long
    Message As Long
    hwnd As Long

End Type

Declare Function CallNextHookEx _
        Lib "user32" (ByVal hHook As Long, _
                      ByVal nCode As Long, _
                      ByVal wParam As Long, _
                      lParam As Any) As Long
Declare Sub CopyMemory _
        Lib "Kernel32" _
        Alias "RtlMoveMemory" (hpvDest As Any, _
                               hpvSource As Any, _
                               ByVal cbCopy As Long)
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowsHookEx _
        Lib "user32" _
        Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                   ByVal lpfn As Long, _
                                   ByVal hmod As Long, _
                                   ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_CALLWNDPROC = 4

Public Const WM_CREATE = &H1

Public hHook As Long

Public Function AppHook(ByVal idHook As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

        On Error GoTo AppHook_Err

        Dim CWP As CWPSTRUCT
100     CopyMemory CWP, ByVal lParam, Len(CWP)

102     Select Case CWP.Message

            Case WM_CREATE
104             SetForegroundWindow CWP.hwnd
106             AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
108             UnhookWindowsHookEx hHook
110             hHook = 0
                Exit Function

        End Select

112     AppHook = CallNextHookEx(hHook, idHook, wParam, ByVal lParam)
        Exit Function
AppHook_Err:
114     Call TraceError(Err.Number, Err.Description, "SysTray.AppHook", Erl)
116

End Function
