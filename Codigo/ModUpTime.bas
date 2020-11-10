Attribute VB_Name = "ModUpTime"

Public SERVER_UPTIME As Long

Public Sub ObtenerUpTime()
        
        On Error GoTo ObtenerUpTime_Err
        

        Dim horas    As Integer

        Dim minutos  As Integer

        Dim segundos As Integer

100     horas = SERVER_UPTIME / 3600
102     minutos = (SERVER_UPTIME Mod 3600) / 60
104     segundos = ((SERVER_UPTIME Mod 3600) Mod 60)

106     MsgBox "Tiempo online: " & horas & " horas; " & minutos & " minutos " & segundos & " segundos."

        
        Exit Sub

ObtenerUpTime_Err:
        Call RegistrarError(Err.Number, Err.description, "ModUpTime.ObtenerUpTime", Erl)
        Resume Next
        
End Sub
