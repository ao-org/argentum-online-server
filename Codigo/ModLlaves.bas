Attribute VB_Name = "ModLlaves"
Option Explicit

' Cantidad máxima de llaves
Public Const MAXKEYS As Byte = 10

Public Function MeterLlaveEnLLavero(ByVal UserIndex As Integer, ByVal Llave As Integer) As Boolean

    On Error GoTo ErrHandler

100     With UserList(UserIndex)

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
110         Call WriteUpdateUserKey(UserIndex, i, Llave)
        
        End With
    
112     MeterLlaveEnLLavero = True
    
        Exit Function

ErrHandler:
114     Call RegistrarError(Err.Number, Err.description, "ModLlaves.MeterLlaveEnLLavero")

End Function

Public Sub SacarLlaveDeLLavero(ByVal UserIndex As Integer, ByVal Llave As Integer)

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
110     Call RegistrarError(Err.Number, Err.description, "ModLlaves.SacarLlaveDeLLavero")

End Sub

Public Sub EnviarLlaves(ByVal UserIndex As Integer)
        
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
        Call RegistrarError(Err.Number, Err.description, "ModLlaves.EnviarLlaves", Erl)

        
End Sub

Public Sub UsarLlave(ByVal UserIndex As Integer, ByVal slot As Integer)
        
        On Error GoTo UsarLlave_Err
    
        

100     If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    
        Dim TargObj As ObjData
        Dim LlaveObj As ObjData
    
102     With UserList(UserIndex)

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
                            
124                             Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                            Else

126                             Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                            End If

                        Else
128                         If TargObj.clave = LlaveObj.clave Then
130                             MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
132                             .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
134                             Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                            Else

136                             Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                            End If

                        End If

                    Else
138                     Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                    End If

                End If
            End If
    
        End With

        
        Exit Sub

UsarLlave_Err:
        Call RegistrarError(Err.Number, Err.description, "ModLlaves.UsarLlave", Erl)

        
End Sub
