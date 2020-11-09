Attribute VB_Name = "GetWrite"
Option Explicit
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
szReturn = ""
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function
Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
writeprivateprofilestring Main, Var, value, file
End Sub
