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
'**************************************************************
' diccionario.cls
'
' Designed and implemented by Mariono Barrou (El Oso)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

'clase diccionario
'basico, plain sin queso ni papa fritas

'mi idea cuando hice esto, lo encontre en el rigido :p. Hecha por el oso

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
        
100     p_cant = 0

        
        Exit Sub

Class_Initialize_Err:
102     Call TraceError(Err.Number, Err.Description, "diccionario.Class_Initialize", Erl)
104
        
End Sub

Public Sub Class_Terminate()

    'destructor
    'destruir los variants?????
End Sub

Public Property Get CantElem() As Integer
        
        On Error GoTo CantElem_Err
        
100     CantElem = p_cant

        
        Exit Property

CantElem_Err:
102     Call TraceError(Err.Number, Err.Description, "diccionario.CantElem", Erl)
104
        
End Property

Public Function AtPut(ByVal clave As String, ByRef elem As Variant) As Boolean
        
        On Error GoTo AtPut_Err
        

        Dim i As Integer

100     AtPut = False

102     If LenB(clave) = 0 Then Exit Function

104     clave = UCase$(clave)

106     If p_cant = MAX_ELEM Then
108         AtPut = False
        Else

110         For i = 1 To p_cant

112             If clave = p_elementos(i).clave Then
114                 p_elementos(i).def = elem
116                 AtPut = True

                    Exit For    ' epa ;)

                End If

118         Next i

120         If Not AtPut Then
122             p_cant = p_cant + 1
124             p_elementos(p_cant).def = elem
126             p_elementos(p_cant).clave = clave
128             AtPut = True

            End If

        End If

        
        Exit Function

AtPut_Err:
130     Call TraceError(Err.Number, Err.Description, "diccionario.AtPut", Erl)
132
        
End Function

Public Function At(ByVal clave As String) As Variant
        
        On Error GoTo At_Err
        

        Dim i As Integer

100     clave = UCase$(clave)

102     At = Null

104     For i = 1 To p_cant

106         If clave = p_elementos(i).clave Then
108             At = p_elementos(i).def
                Exit Function

            End If

110     Next i

        
        Exit Function

At_Err:
112     Call TraceError(Err.Number, Err.Description, "diccionario.At", Erl)
114
        
End Function

Public Function AtIndex(ByVal i As Integer) As String
        
        On Error GoTo AtIndex_Err
        
100     AtIndex = p_elementos(i).clave

        
        Exit Function

AtIndex_Err:
102     Call TraceError(Err.Number, Err.Description, "diccionario.AtIndex", Erl)
104
        
End Function

Public Function MayorValor(ByRef cant As Integer) As String
        
        On Error GoTo MayorValor_Err
        

        'parchecito para el AO, me da la clave con mayor valor en valor
        'y la cantidad de claves con ese valor (por si hay empate)
        Dim i     As Integer

        Dim max   As Integer

        Dim clave As String

100     max = -1
102     cant = 0
104     clave = vbNullString

106     For i = 1 To p_cant

108         If max <= CInt(p_elementos(i).def) Then
110             cant = IIf(max = CInt(p_elementos(i).def), cant + 1, 1)
112             clave = IIf(max = CInt(p_elementos(i).def), clave & "," & p_elementos(i).clave, p_elementos(i).clave)
114             max = CInt(p_elementos(i).def)

            End If

116     Next i

118     MayorValor = clave

        
        Exit Function

MayorValor_Err:
120     Call TraceError(Err.Number, Err.Description, "diccionario.MayorValor", Erl)
122
        
End Function

Public Sub DumpAll()
        
        On Error GoTo DumpAll_Err
        

        Dim i As Integer

100     For i = 1 To MAX_ELEM
102         p_elementos(i).clave = vbNullString
104         p_elementos(i).def = Null
106     Next i

108     p_cant = 0

        
        Exit Sub

DumpAll_Err:
110     Call TraceError(Err.Number, Err.Description, "diccionario.DumpAll", Erl)
112
        
End Sub
