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
    On Error Goto SacarLlaveDeLLavero_Err

        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)

            Dim i As Integer
            
102         For i = 1 To MAXKEYS
104             If .Keys(i) = Llave Then
106                 .Keys(i) = 0
108                 Call WriteUpdateUserKey(UserIndex, i, 0)
                    Exit Sub
                End If
            Next
    
        End With
    
        Exit Sub

ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "ModLlaves.SacarLlaveDeLLavero", Erl)

    Exit Sub
SacarLlaveDeLLavero_Err:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.SacarLlaveDeLLavero", Erl)
End Sub

Public Sub EnviarLlaves(ByVal UserIndex As Integer)
    On Error Goto EnviarLlaves_Err
        
        On Error GoTo EnviarLlaves_Err
    
        
    
100     With UserList(UserIndex)

            Dim i As Integer
            
102         For i = 1 To MAXKEYS
104             If .Keys(i) <> 0 Then
106                 Call WriteUpdateUserKey(UserIndex, i, .Keys(i))
                End If
            Next
    
        End With
        
        Exit Sub

EnviarLlaves_Err:
108     Call TraceError(Err.Number, Err.Description, "ModLlaves.EnviarLlaves", Erl)

        
    Exit Sub
EnviarLlaves_Err:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.EnviarLlaves", Erl)
End Sub

Public Sub UsarLlave(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error Goto UsarLlave_Err
        
        On Error GoTo UsarLlave_Err
    
        

100     If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    
        Dim TargObj As t_ObjData
        Dim LlaveObj As t_ObjData
    
102     With UserList(UserIndex)

            If Slot > MAXKEYS Then
                'Call BanearIP(0, UserList(UserIndex).name, UserList(UserIndex).IP, UserList(UserIndex).Cuenta)
                Call LogEdicionPaquete("El usuario " & UserList(UserIndex).Name & " editó el slot del llavero | Valor: " & Slot & ".")
                Exit Sub
            End If
104         If .Keys(Slot) <> 0 Then
106             If .flags.TargetObj = 0 Then Exit Sub
            
108             TargObj = ObjData(.flags.TargetObj)
110             LlaveObj = ObjData(.Keys(Slot))

                '¿El objeto clickeado es una puerta?
112             If TargObj.OBJType = e_OBJType.otPuertas Then

                    '¿Esta cerrada?
114                 If TargObj.Cerrada = 1 Then

                        '¿Cerrada con llave?
116                     If TargObj.Llave > 0 Then
118                         If TargObj.clave = LlaveObj.clave Then 'Or LlaveObj.clave = "3450" Then
120                             MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
122                             .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
124                             Call WriteLocaleMsg(UserIndex, "897", e_FontTypeNames.FONTTYPE_INFO)
                            Else
126                             Call WriteLocaleMsg(UserIndex, "898", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        Else
128                         If TargObj.clave = LlaveObj.clave Then 'Or LlaveObj.clave = "3450" Then
130                             MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
132                             .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
134                             'Msg899= Has cerrado con llave la puerta.
                                Call WriteLocaleMsg(UserIndex, "899", e_FontTypeNames.FONTTYPE_INFO)
                            Else

136                             Call WriteLocaleMsg(UserIndex, "898", e_FontTypeNames.FONTTYPE_INFO)
                            End If

                        End If

                    Else
138                     'Msg901= No esta cerrada.
                        Call WriteLocaleMsg(UserIndex, "901", e_FontTypeNames.FONTTYPE_INFO)
                    End If

                End If
            End If
    
        End With

        
        Exit Sub

UsarLlave_Err:
140     Call TraceError(Err.Number, Err.Description, "ModLlaves.UsarLlave", Erl)

        
    Exit Sub
UsarLlave_Err:
    Call TraceError(Err.Number, Err.Description, "ModLlaves.UsarLlave", Erl)
End Sub
