Attribute VB_Name = "DatabaseGlobals"
Option Explicit

Option Base 0

Public Database_Enabled     As Boolean
Public Database_Driver      As String
Public Database_Source      As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Public Const MAX_ASYNC     As Byte = 20
Public Current_async       As Byte

Public Connection          As ADODB.Connection
Public Connection_async(1 To MAX_ASYNC)    As ADODB.Connection

Public Builder             As cStringBuilder

Public Function post_increment(ByRef value As Integer) As Integer
     post_increment = value
     value = value + 1
End Function
