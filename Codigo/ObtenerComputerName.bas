Attribute VB_Name = "ObtenerComputerName"

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal LpBuffer As String, nsize As Long) As Long
 
Public Function ComputerName() As String
    '-- Funcion auxiliar que devuelve el nombre del equipo llamando al API
    ComputerName = Space$(260)
    GetComputerName ComputerName, Len(ComputerName)
    ComputerName = Left$(ComputerName, InStr(ComputerName, vbNullChar) - 1)

End Function
