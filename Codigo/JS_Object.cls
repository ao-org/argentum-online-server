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
100     Set StringBuilder = New cStringBuilder
End Sub

Private Sub Class_Terminate()
100     Set StringBuilder = Nothing
End Sub

Public Sub Clear()
100     StringBuilder.Clear
End Sub

Public Property Let Item(ByRef Key As String, ByRef Value As Variant)
100     StringBuilder.Append """" & Key & """: "
    
102     Select Case VarType(Value)
            Case vbString
104             StringBuilder.Append """" & Value & """"
            
106         Case vbBoolean
108             StringBuilder.Append IIf(Value, "true", "false")
            
110         Case vbObject
112             StringBuilder.Append Value.ToString
            
114         Case Else
116             StringBuilder.Append CStr(Value)
        End Select
    
118     StringBuilder.Append ", "
End Property

Public Function ToString() As String
100     ToString = "{" & Left$(StringBuilder.ToString, Len(StringBuilder.ToString) - 2) & "}"
End Function
