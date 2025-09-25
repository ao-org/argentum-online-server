Attribute VB_Name = "ModLlaves"
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
' Cantidad máxima de llaves
Public Const MAXKEYS As Byte = 10

Public Sub SacarLlaveDeLLavero(ByVal UserIndex As Integer, ByVal Llave As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To MAXKEYS
            If .Keys(i) = Llave Then
                .Keys(i) = 0
                Call WriteUpdateUserKey(UserIndex, i, 0)
                Exit Sub
            End If
        Next
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.SacarLlaveDeLLavero", Erl)
End Sub

Public Sub EnviarLlaves(ByVal UserIndex As Integer)
    On Error GoTo EnviarLlaves_Err
    With UserList(UserIndex)
        Dim i As Integer
        For i = 1 To MAXKEYS
            If .Keys(i) <> 0 Then
                Call WriteUpdateUserKey(UserIndex, i, .Keys(i))
            End If
        Next
    End With
    Exit Sub
EnviarLlaves_Err:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.EnviarLlaves", Erl)
End Sub

Public Sub UsarLlave(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error GoTo UsarLlave_Err
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    Dim TargObj  As t_ObjData
    Dim LlaveObj As t_ObjData
    With UserList(UserIndex)
        If Slot > MAXKEYS Then
            'Call BanearIP(0, UserList(UserIndex).name, UserList(UserIndex).IP, UserList(UserIndex).Cuenta)
            Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " editó el slot del llavero | Valor: " & Slot & ".")
            Exit Sub
        End If
        If .Keys(Slot) <> 0 Then
            If .flags.TargetObj = 0 Then Exit Sub
            TargObj = ObjData(.flags.TargetObj)
            LlaveObj = ObjData(.Keys(Slot))
            '¿El objeto clickeado es una puerta?
            If TargObj.OBJType = e_OBJType.otDoors Then
                '¿Esta cerrada?
                If TargObj.Cerrada = 1 Then
                    '¿Cerrada con llave?
                    If TargObj.Llave > 0 Then
                        If TargObj.clave = LlaveObj.clave Then 'Or LlaveObj.clave = "3450" Then
                            MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, _
                                    .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                            .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            Call WriteLocaleMsg(UserIndex, "897", e_FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteLocaleMsg(UserIndex, "898", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        If TargObj.clave = LlaveObj.clave Then 'Or LlaveObj.clave = "3450" Then
                            MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, _
                                    .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                            .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            'Msg899= Has cerrado con llave la puerta.
                            Call WriteLocaleMsg(UserIndex, "899", e_FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteLocaleMsg(UserIndex, "898", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Else
                    'Msg901= No esta cerrada.
                    Call WriteLocaleMsg(UserIndex, "901", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    Exit Sub
UsarLlave_Err:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.UsarLlave", Erl)
End Sub
