Attribute VB_Name = "modInvisibles"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

' 0 = viejo
' 1 = nuevo
#Const MODO_INVISIBILIDAD = 0

' cambia el estado de invisibilidad a 1 o 0 dependiendo del modo: true o false
'
Public Sub PonerInvisible(ByVal UserIndex As Integer, ByVal estado As Boolean)
        
        On Error GoTo PonerInvisible_Err
        
        #If MODO_INVISIBILIDAD = 0 Then

100         UserList(UserIndex).flags.invisible = IIf(estado, 1, 0)
102         UserList(UserIndex).flags.Oculto = IIf(estado, 1, 0)
104         UserList(UserIndex).Counters.Invisibilidad = 0

106         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.charindex, Not Estado, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))

        #Else

            Dim EstadoActual As Boolean

            ' Está invisible ?
108         EstadoActual = (UserList(UserIndex).flags.invisible = 1)

            'If EstadoActual <> Modo Then
110         If Modo = True Then
                ' Cuando se hace INVISIBLE se les envia a los
                ' clientes un Borrar Char
112             UserList(UserIndex).flags.invisible = 1
                '        'Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",1")
114             Call SendData(SendTarget.toMap, UserList(UserIndex).Pos.Map, PrepareMessageCharacterRemove(1, UserList(UserIndex).Char.CharIndex, True))
            Else
        
            End If

            'End If

        #End If

        
        Exit Sub

PonerInvisible_Err:
116     Call TraceError(Err.Number, Err.Description, "modInvisibles.PonerInvisible", Erl)

        
End Sub

