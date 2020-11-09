Attribute VB_Name = "ModUpTime"

Public SERVER_UPTIME As Long

Public Sub ObtenerUpTime()

    Dim horas    As Integer

    Dim minutos  As Integer

    Dim segundos As Integer

    horas = SERVER_UPTIME / 3600
    minutos = (SERVER_UPTIME Mod 3600) / 60
    segundos = ((SERVER_UPTIME Mod 3600) Mod 60)

    MsgBox "Tiempo online: " & horas & " horas; " & minutos & " minutos " & segundos & " segundos."

End Sub
