VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JS_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StringBuilder As cStringBuilder

Private Sub Class_Initialize()
    Set StringBuilder = New cStringBuilder
End Sub

Private Sub Class_Terminate()
    Set StringBuilder = Nothing
End Sub

Public Sub Clear()
    StringBuilder.Clear
End Sub

Public Property Let Item(ByRef Key As String, ByRef Value As Variant)
    StringBuilder.Append """" & Key & """: "
    
    Select Case VarType(Value)
        Case vbString
            StringBuilder.Append """" & Value & """"
            
        Case vbBoolean
            StringBuilder.Append IIf(Value, "true", "false")
            
        Case vbObject
            StringBuilder.Append Value.ToString
            
        Case Else
            StringBuilder.Append CStr(Value)
            
    End Select
    
    StringBuilder.Append ", "
End Property

Public Function ToString() As String
    ToString = "{" & Left$(StringBuilder.ToString, Len(StringBuilder.ToString) - 2) & "}"
End Function

