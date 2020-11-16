Attribute VB_Name = "ModLlaves"
Option Explicit

' Cantidad máxima de llaves
Public Const MAXKEYS As Byte = 10

Public Function MeterLlaveEnLLavero(ByVal UserIndex As Integer, ByVal Llave As Integer) As Boolean

On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim i As Integer
        
        For i = 1 To MAXKEYS
            If .Keys(i) = 0 Then
                Exit For
            End If
        Next
        
        ' No hay espacio
        If i > MAXKEYS Then Exit Function
        
        ' Metemos la llave
        .Keys(i) = Llave
        Call WriteUpdateUserKey(UserIndex, i, Llave)
        
    End With
    
    MeterLlaveEnLLavero = True
    
    Exit Function

Errhandler:
    Call RegistrarError(Err.Number, Err.description, "ModLlaves.MeterLlaveEnLLavero")

End Function

Public Sub SacarLlaveDeLLavero(ByVal UserIndex As Integer, ByVal Llave As Integer)

    On Error GoTo Errhandler
    
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

Errhandler:
    Call RegistrarError(Err.Number, Err.description, "ModLlaves.SacarLlaveDeLLavero")

End Sub

Public Sub EnviarLlaves(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)

        Dim i As Integer
            
        For i = 1 To MAXKEYS
            If .Keys(i) <> 0 Then
                Call WriteUpdateUserKey(UserIndex, i, .Keys(i))
            End If
        Next
    
    End With
End Sub

Public Sub UsarLlave(ByVal UserIndex As Integer, ByVal slot As Integer)

    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    
    Dim TargObj As ObjData
    Dim LlaveObj As ObjData
    
    With UserList(UserIndex)

        If .Keys(slot) <> 0 Then
            If .flags.TargetObj = 0 Then Exit Sub
            
            TargObj = ObjData(.flags.TargetObj)
            LlaveObj = ObjData(.Keys(slot))

            '¿El objeto clickeado es una puerta?
            If TargObj.OBJType = eOBJType.otPuertas Then

                '¿Esta cerrada?
                If TargObj.Cerrada = 1 Then

                    '¿Cerrada con llave?
                    If TargObj.Llave > 0 Then
                        If TargObj.clave = LlaveObj.clave Then
                            MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                            .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
                            Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        End If

                    Else
                        If TargObj.clave = LlaveObj.clave Then
                            MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                            .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                            
                            Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        End If

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                End If

            End If
        End If
    
    End With

End Sub
