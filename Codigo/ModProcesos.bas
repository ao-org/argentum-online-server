Attribute VB_Name = "ModProcesos"
Option Explicit

Private dcnProcesosExcluidos As Dictionary

Public Function OrderProcesses(ByRef data As String)
    On Error GoTo OrderProcesses_Err
    
    Dim listProcess() As String
    Dim processListNotRepeated As Dictionary
    Dim i As Integer, j As Integer, k As Integer
    
    Set processListNotRepeated = New Dictionary
    Call CargarProcesosExcluidos
    
    'Convierto la data a array() string
    listProcess = Split(data, vbNewLine)
    
    For i = 0 To UBound(listProcess)
        If processListNotRepeated.Exists(listProcess(i)) Then
            processListNotRepeated(listProcess(i)) = processListNotRepeated(listProcess(i)) + 1
        Else
            Call processListNotRepeated.Add(listProcess(i), 1)
        End If
    Next i
    
    data = ""
    Dim keys() As String
    keys = processListNotRepeated.keys
    For i = 0 To processListNotRepeated.Count - 1
        If Not dcnProcesosExcluidos.Exists(keys(i)) Then
            data = data & "(?)"
        End If
        data = data & keys(i) & "(" & processListNotRepeated(keys(i)) & ")" & vbNewLine
    Next
    
    Debug.Print ""
        
   

  
OrderProcesses_Err:
120     Call RegistrarError(Err.Number, Err.Description, "ES.OrderProcesses", Erl)
122     Resume Next
          
    
    

    
End Function

Public Sub CargarProcesosExcluidos()
        
    On Error GoTo CargarProcesosExcluidos_Err
    Dim n As Integer, i As Integer, cad As String
    
    Set dcnProcesosExcluidos = New Dictionary
108 n = FreeFile(1)

110 Open DatPath & "procesos_excluidos.dat" For Input As #n
    
106 Do While Not EOF(n)
110     Line Input #n, cad
        dcnProcesosExcluidos(cad) = 1
    Loop
    
112 Close n
      
    Exit Sub

CargarProcesosExcluidos_Err:
120     Call RegistrarError(Err.Number, Err.Description, "ES.CargarProcesosExcluidos", Erl)
122     Resume Next
        
End Sub

