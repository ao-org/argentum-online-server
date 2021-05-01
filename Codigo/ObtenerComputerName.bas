Attribute VB_Name = "ObtenerComputerName"

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal LpBuffer As String, nsize As Long) As Long
 
Public Function ComputerName() As String
        '-- Funcion auxiliar que devuelve el nombre del equipo llamando al API
        
        On Error GoTo ComputerName_Err
        
100     ComputerName = Space$(260)
102     GetComputerName ComputerName, Len(ComputerName)
104     ComputerName = Left$(ComputerName, InStr(ComputerName, vbNullChar) - 1)

        
        Exit Function

ComputerName_Err:
106     Call RegistrarError(Err.Number, Err.Description, "ObtenerComputerName.ComputerName", Erl)
108     Resume Next
        
End Function
