Attribute VB_Name = "ModLlaves"
Option Explicit

' Cantidad máxima de llaves
Public Const MAXKEYS As Byte = 10

Public Function MeterLlaveEnLLavero(ByVal Userindex As Integer, ByVal Llave As Integer) As Boolean

    On Error GoTo ErrHandler

100     With UserList(Userindex)

            Dim i As Integer
        
102         For i = 1 To MAXKEYS
104             If .Keys(i) = 0 Then
                    Exit For
                End If
            Next
        
            ' No hay espacio
106         If i > MAXKEYS Then Exit Function
        
            ' Metemos la llave
108         .Keys(i) = Llave
110         Call WriteUpdateUserKey(Userindex, i, Llave)
        
        End With
    
112     MeterLlaveEnLLavero = True
    
        Exit Function

ErrHandler:
114     Call RegistrarError(Err.Number, Err.description, "ModLlaves.MeterLlaveEnLLavero")

End Function

Public Sub SacarLlaveDeLLavero(ByVal Userindex As Integer, ByVal Llave As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(Userindex)

            Dim i As Integer
            
102         For i = 1 To MAXKEYS
104             If .Keys(i) = Llave Then
106                 .Keys(i) = 0
108                 Call WriteUpdateUserKey(Userindex, i, 0)
                    Exit Sub
                End If
            Next
    
        End With
    
        Exit Sub

ErrHandler:
110     Call RegistrarError(Err.Number, Err.description, "ModLlaves.SacarLlaveDeLLavero")

End Sub

Public Sub EnviarLlaves(ByVal Userindex As Integer)
    
100     With UserList(Userindex)

            Dim i As Integer
            
102         For i = 1 To MAXKEYS
104             If .Keys(i) <> 0 Then
106                 Call WriteUpdateUserKey(Userindex, i, .Keys(i))
                End If
            Next
    
        End With
End Sub

Public Sub UsarLlave(ByVal Userindex As Integer, ByVal slot As Integer)

100     If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
    
        Dim TargObj As ObjData
        Dim LlaveObj As ObjData
    
102     With UserList(Userindex)

104         If .Keys(slot) <> 0 Then
106             If .flags.TargetObj = 0 Then Exit Sub
            
108             TargObj = ObjData(.flags.TargetObj)
110             LlaveObj = ObjData(.Keys(slot))

                '¿El objeto clickeado es una puerta?
112             If TargObj.OBJType = eOBJType.otPuertas Then

                    '¿Esta cerrada?
114                 If TargObj.Cerrada = 1 Then

                        '¿Cerrada con llave?
116                     If TargObj.Llave > 0 Then
118                         If TargObj.clave = LlaveObj.clave Then
120                             MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
122                             .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
124                             Call WriteConsoleMsg(Userindex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                            Else

126                             Call WriteConsoleMsg(Userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                            End If

                        Else
128                         If TargObj.clave = LlaveObj.clave Then
130                             MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
132                             .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
134                             Call WriteConsoleMsg(Userindex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                            Else

136                             Call WriteConsoleMsg(Userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                            End If

                        End If

                    Else
138                     Call WriteConsoleMsg(Userindex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                    End If

                End If
            End If
    
        End With

End Sub
