VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "diccionario"
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
Private Const MAX_ELEM = 100

Private Type t_diccElem
    clave As String
    def As Variant
End Type

Private p_elementos(1 To MAX_ELEM) As t_diccElem 'visual basic es una mierda para usar memoria dinamica, asi que uso esto
Private p_cant                     As Integer

Public Sub Class_Initialize()
    'constructor
    On Error GoTo Class_Initialize_Err
    p_cant = 0
    Exit Sub
Class_Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.Class_Initialize", Erl)
End Sub

Public Sub Class_Terminate()
    'destructor
    'destruir los variants?????
End Sub

Public Property Get CantElem() As Integer
    On Error GoTo CantElem_Err
    CantElem = p_cant
    Exit Property
CantElem_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.CantElem", Erl)
End Property

Public Function AtPut(ByVal clave As String, ByRef elem As Variant) As Boolean
    On Error GoTo AtPut_Err
    Dim i As Integer
    AtPut = False
    If LenB(clave) = 0 Then Exit Function
    clave = UCase$(clave)
    If p_cant = MAX_ELEM Then
        AtPut = False
    Else
        For i = 1 To p_cant
            If clave = p_elementos(i).clave Then
                p_elementos(i).def = elem
                AtPut = True
                Exit For    ' epa ;)
            End If
        Next i
        If Not AtPut Then
            p_cant = p_cant + 1
            p_elementos(p_cant).def = elem
            p_elementos(p_cant).clave = clave
            AtPut = True
        End If
    End If
    Exit Function
AtPut_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.AtPut", Erl)
End Function

Public Function At(ByVal clave As String) As Variant
    On Error GoTo At_Err
    Dim i As Integer
    clave = UCase$(clave)
    At = Null
    For i = 1 To p_cant
        If clave = p_elementos(i).clave Then
            At = p_elementos(i).def
            Exit Function
        End If
    Next i
    Exit Function
At_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.At", Erl)
End Function

Public Function AtIndex(ByVal i As Integer) As String
    On Error GoTo AtIndex_Err
    AtIndex = p_elementos(i).clave
    Exit Function
AtIndex_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.AtIndex", Erl)
End Function

Public Function MayorValor(ByRef cant As Integer) As String
    On Error GoTo MayorValor_Err
    'parchecito para el AO, me da la clave con mayor valor en valor
    'y la cantidad de claves con ese valor (por si hay empate)
    Dim i     As Integer
    Dim max   As Integer
    Dim clave As String
    max = -1
    cant = 0
    clave = vbNullString
    For i = 1 To p_cant
        If max <= CInt(p_elementos(i).def) Then
            cant = IIf(max = CInt(p_elementos(i).def), cant + 1, 1)
            clave = IIf(max = CInt(p_elementos(i).def), clave & "," & p_elementos(i).clave, p_elementos(i).clave)
            max = CInt(p_elementos(i).def)
        End If
    Next i
    MayorValor = clave
    Exit Function
MayorValor_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.MayorValor", Erl)
End Function

Public Sub DumpAll()
    On Error GoTo DumpAll_Err
    Dim i As Integer
    For i = 1 To MAX_ELEM
        p_elementos(i).clave = vbNullString
        p_elementos(i).def = Null
    Next i
    p_cant = 0
    Exit Sub
DumpAll_Err:
    Call TraceError(Err.Number, Err.Description, "diccionario.DumpAll", Erl)
End Sub
