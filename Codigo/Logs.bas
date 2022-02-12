Attribute VB_Name = "Logs"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Public LogsBuffer As New cStringBuilder
Private Const MAX_LOG_SIZE As Long = 1000000 ' 1MB en el buffer antes de volcarlo al .log

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type: Private HistorialError As UltimoError


Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************

    On Error GoTo RegistrarError_Err
    
    'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
    If Componente = HistorialError.Componente And _
       Numero = HistorialError.ErrorCode Then
       
       'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
        'x lo que no hace falta registrar el error.
        If HistorialError.Contador = 10 Then
            'Debug.Assert False
            Exit Sub
        End If
        
        'Agregamos el error al historial.
        HistorialError.Contador = HistorialError.Contador + 1
        
    Else 'Si NO es igual, reestablecemos el contador.

        HistorialError.Contador = 0
        HistorialError.ErrorCode = Numero
        HistorialError.Componente = Componente
            
    End If
    
    ' ----------------------------------------------------------------------------------
    ' Jopi: Guardamos los errores en un String Buffer
    '  para no andar haciendo operaciones I/O (costosas) cada vez que entra un error
    ' ----------------------------------------------------------------------------------
    LogsBuffer.AppendNL "Error: " & Numero
    LogsBuffer.AppendNL "Descripcion: " & Descripcion
    LogsBuffer.AppendNL "Componente: " & Componente
    
    If LenB(Linea) <> 0 Then
        LogsBuffer.AppendNL "Linea: " & Linea
    End If
    
    LogsBuffer.AppendNL "Fecha y Hora: " & Date$ & "-" & Time$
    
    LogsBuffer.AppendNL vbNullString
    
    ' ----------------------------------------------------------------------------------------------
    ' Jopi: Una vez que el buffer llega a cierta capacidad, volcamos los contenidos al archivo .log
    ' ----------------------------------------------------------------------------------------------
    'If LogsBuffer.ByteLength > MAX_LOG_SIZE Then
        Dim File As Integer: File = FreeFile
        
        Open App.Path & "\logs\Errores\General.log" For Append As #File
            Print #File, LogsBuffer.ToString
        Close #File
        
        ' Limpiamos el buffer
        Call LogsBuffer.Clear
    'End If
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Linea: " & Linea & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
        
    Exit Sub
    
RegistrarError_Err:
    Close #File
        
End Sub

