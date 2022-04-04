VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
'                    Metodos publicos
'
' Public sub Push(byval i as variant) mete el elemento i
' al final de la cola.
'
' Public Function Pop As Variant: quita de la cola el primer elem
' y lo devuelve
'
' Public Function VerElemento(ByVal Index As Integer) As Variant
' muestra el elemento numero Index de la cola sin quitarlo
'
' Public Function PopByVal() As Variant: muestra el primer
' elemento de la cola sin quitarlo
'
' Public Property Get Longitud() As Integer: devuelve la
' cantidad de elementos que tiene la cola.

Private Const FRENTE = 1

Private Cola As Collection

Public Sub Reset()
        
        On Error GoTo Reset_Err
    
        

        

        Dim i As Integer

100     For i = 1 To Me.Longitud
102         Cola.Remove FRENTE
104     Next i

        
        Exit Sub

Reset_Err:
106     Call TraceError(Err.Number, Err.Description, "cCola.Reset", Erl)

        
End Sub

Public Property Get Longitud() As Integer
        
        On Error GoTo Longitud_Err
        
100     Longitud = Cola.Count

        
        Exit Property

Longitud_Err:
102     Call TraceError(Err.Number, Err.Description, "cCola.Longitud", Erl)
104
        
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
        
        On Error GoTo IndexValido_Err
        
100     IndexValido = i >= 1 And i <= Me.Longitud

        
        Exit Function

IndexValido_Err:
102     Call TraceError(Err.Number, Err.Description, "cCola.IndexValido", Erl)
104
        
End Function

Private Sub Class_Initialize()
        
        On Error GoTo Class_Initialize_Err
        
100     Set Cola = New Collection

        
        Exit Sub

Class_Initialize_Err:
102     Call TraceError(Err.Number, Err.Description, "cCola.Class_Initialize", Erl)
104
        
End Sub

Public Function VerElemento(ByVal Index As Integer) As String
        
        On Error GoTo VerElemento_Err
    
        

        

100     If IndexValido(Index) Then
            'Pablo
            'VerElemento = UCase$(Cola.Item(index))
            '/Pablo
102         VerElemento = Cola(Index)
        Else
104         VerElemento = 0

        End If

        
        Exit Function

VerElemento_Err:
106     Call TraceError(Err.Number, Err.Description, "cCola.VerElemento", Erl)

        
End Function

Public Sub Push(ByVal nombre As String, ByVal Consulta As String, ByVal TipoDeConsulta As String)
        
        On Error GoTo Push_Err
    
        

        

        'Mete elemento en la cola
100     Call Cola.Add(nombre & "Ø" & Consulta & "Ø" & TipoDeConsulta)

        
        Exit Sub

Push_Err:
102     Call TraceError(Err.Number, Err.Description, "cCola.Push", Erl)

        
End Sub

Public Function Pop() As String
        
        On Error GoTo Pop_Err
    
        

        

        'Quita elemento de la cola
100     If Cola.Count > 0 Then
102         Pop = Cola(FRENTE)
104         Call Cola.Remove(FRENTE)
        Else
106         Pop = 0

        End If

        
        Exit Function

Pop_Err:
108     Call TraceError(Err.Number, Err.Description, "cCola.Pop", Erl)

        
End Function

Public Function PopByVal() As String
        
        On Error GoTo PopByVal_Err

        'Quita elemento de la cola
100     If Cola.Count > 0 Then
102         PopByVal = Cola.Item(1)
        Else
104         PopByVal = 0

        End If

        
        Exit Function

PopByVal_Err:
106     Call TraceError(Err.Number, Err.Description, "cCola.PopByVal", Erl)

        
End Function

Public Function Existe(ByVal nombre As String) As Boolean
        
        On Error GoTo Existe_Err
    
        

        
  
        Dim V                  As String

        Dim i                  As Integer

        Dim NombreEnMayusculas As String

100     NombreEnMayusculas = UCase$(nombre)
 
102     For i = 1 To Me.Longitud
104         V = UCase$(ReadField(1, Me.VerElemento(i), Asc("Ø")))
106         If V = NombreEnMayusculas Then
108             Existe = True
                Exit Function

            End If

        Next
110     Existe = False

        
        Exit Function

Existe_Err:
112     Call TraceError(Err.Number, Err.Description, "cCola.Existe", Erl)

        
End Function

Public Sub Quitar(ByVal nombre As String)
        
        On Error GoTo Quitar_Err
    
        

        

        Dim V                  As String

        Dim i                  As Integer

        Dim NombreEnMayusculas As String
 
        'NombreEnMayusculas = UCase$(Nombre)
 
100     For i = 1 To Me.Longitud
            'Pablo
            'V = mid$(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
            '/Pablo
102         V = Me.VerElemento(i)

104         If V = nombre Then
106             Call Cola.Remove(i)
                Exit Sub

            End If

108     Next i
 
        
        Exit Sub

Quitar_Err:
110     Call TraceError(Err.Number, Err.Description, "cCola.Quitar", Erl)

        
End Sub

Public Sub QuitarIndex(ByVal Index As Integer)
        
        On Error GoTo QuitarIndex_Err
    
        

        

100     If IndexValido(Index) Then Call Cola.Remove(Index)

        
        Exit Sub

QuitarIndex_Err:
102     Call TraceError(Err.Number, Err.Description, "cCola.QuitarIndex", Erl)

        
End Sub

Private Sub Class_Terminate()
        'Destruimos el objeto Cola
        
        On Error GoTo Class_Terminate_Err
        
100     Set Cola = Nothing

        
        Exit Sub

Class_Terminate_Err:
102     Call TraceError(Err.Number, Err.Description, "cCola.Class_Terminate", Erl)
104
        
End Sub
