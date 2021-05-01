VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JS_Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StringBuilder As cStringBuilder

Private Sub Class_Initialize()
100     Set StringBuilder = New cStringBuilder
End Sub

Public Sub Clear()
100     StringBuilder.Clear
End Sub

Public Sub Push(ByRef Value As Variant)
100     Select Case VarType(Value)
            Case vbString
102             StringBuilder.Append """" & Value & """"
104         Case vbBoolean
106             StringBuilder.Append IIf(Value, "true", "false")
108         Case vbObject
110             StringBuilder.Append Value.ToString
112         Case Else
114             StringBuilder.Append CStr(Value)
        End Select
    
116     StringBuilder.Append ", "
End Sub

Public Function ToString() As String
100     ToString = "[" & Left$(StringBuilder.ToString, Len(StringBuilder.ToString) - 2) & "]"
End Function
