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
    Set StringBuilder = New cStringBuilder
End Sub

Public Sub Clear()
    StringBuilder.Clear
End Sub

Public Sub Push(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbString
            StringBuilder.Append """" & Value & """"
        Case vbBoolean
            StringBuilder.Append IIf(Value, "true", "false")
        Case vbNull
            StringBuilder.Append "null"
        Case vbObject
            StringBuilder.Append Value.ToString
        Case Else
            StringBuilder.Append CStr(Value)
    End Select
    
    StringBuilder.Append ", "
End Sub

Public Function ToString() As String
    If StringBuilder.Length > 0 Then
        ToString = "[" & Left$(StringBuilder.ToString, StringBuilder.Length - 2) & "]"
    Else
        ToString = "[]"
    End If
End Function

