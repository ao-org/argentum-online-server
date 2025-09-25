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
    For i = 1 To Me.Longitud
        Cola.Remove FRENTE
    Next i
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Reset", Erl)
End Sub

Public Property Get Longitud() As Integer
    On Error GoTo Longitud_Err
    Longitud = Cola.count
    Exit Property
Longitud_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Longitud", Erl)
End Property

Private Function IndexValido(ByVal i As Integer) As Boolean
    On Error GoTo IndexValido_Err
    IndexValido = i >= 1 And i <= Me.Longitud
    Exit Function
IndexValido_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.IndexValido", Erl)
End Function

Private Sub Class_Initialize()
    On Error GoTo Class_Initialize_Err
    Set Cola = New Collection
    Exit Sub
Class_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Class_Initialize", Erl)
End Sub

Public Function VerElemento(ByVal Index As Integer) As String
    On Error GoTo VerElemento_Err
    If IndexValido(Index) Then
        VerElemento = Cola(Index)
    Else
        VerElemento = 0
    End If
    Exit Function
VerElemento_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.VerElemento", Erl)
End Function

Public Sub Push(ByVal nombre As String, ByVal Consulta As String, ByVal TipoDeConsulta As String)
    On Error GoTo Push_Err
    Dim FechaHoraDeConsulta As Date
    FechaHoraDeConsulta = Now
    'Mete elemento en la cola
    Call Cola.Add(nombre & "Ø" & Consulta & "Ø" & TipoDeConsulta & "Ø" & FechaHoraDeConsulta)
    Exit Sub
Push_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Push", Erl)
End Sub

Public Function Pop() As String
    On Error GoTo Pop_Err
    'Quita elemento de la cola
    If Cola.count > 0 Then
        Pop = Cola(FRENTE)
        Call Cola.Remove(FRENTE)
    Else
        Pop = 0
    End If
    Exit Function
Pop_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Pop", Erl)
End Function

Public Function PopByVal() As String
    On Error GoTo PopByVal_Err
    'Quita elemento de la cola
    If Cola.count > 0 Then
        PopByVal = Cola.Item(1)
    Else
        PopByVal = 0
    End If
    Exit Function
PopByVal_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.PopByVal", Erl)
End Function

Public Function Existe(ByVal nombre As String) As Boolean
    On Error GoTo Existe_Err
    Dim v                  As String
    Dim i                  As Integer
    Dim NombreEnMayusculas As String
    NombreEnMayusculas = UCase$(nombre)
    For i = 1 To Me.Longitud
        v = UCase$(ReadField(1, Me.VerElemento(i), Asc("Ø")))
        If v = NombreEnMayusculas Then
            Existe = True
            Exit Function
        End If
    Next
    Existe = False
    Exit Function
Existe_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Existe", Erl)
End Function

Public Sub Quitar(ByVal nombre As String)
    On Error GoTo Quitar_Err
    Dim mensajeEncoladoCompleto   As String
    Dim mensajeSinTimestamp       As String
    Dim posicionUltimoDelimitador As Integer
    Dim i                         As Integer
    For i = 1 To Me.Longitud
        mensajeEncoladoCompleto = Me.VerElemento(i)
        posicionUltimoDelimitador = InStrRev(mensajeEncoladoCompleto, "Ø")
        mensajeSinTimestamp = Left(mensajeEncoladoCompleto, posicionUltimoDelimitador - 1)
        If mensajeSinTimestamp = nombre Then
            Call Cola.Remove(i)
            Exit Sub
        End If
    Next i
    Exit Sub
Quitar_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Quitar", Erl)
End Sub

Public Sub QuitarIndex(ByVal Index As Integer)
    On Error GoTo QuitarIndex_Err
    If IndexValido(Index) Then Call Cola.Remove(Index)
    Exit Sub
QuitarIndex_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.QuitarIndex", Erl)
End Sub

Private Sub Class_Terminate()
    'Destruimos el objeto Cola
    On Error GoTo Class_Terminate_Err
    Set Cola = Nothing
    Exit Sub
Class_Terminate_Err:
    Call TraceError(Err.Number, Err.Description, "cCola.Class_Terminate", Erl)
End Sub
