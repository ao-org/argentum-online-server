VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "TrapConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function GetValue(ByVal Name As String) As Long
    Select Case Name
        Case "TrapPlacementRange": GetValue = 4
        Case "TrapParaliceDuration": GetValue = 2
        Case Else: Err.Raise 5, , "Unexpected setting: " & Name
    End Select
End Function
