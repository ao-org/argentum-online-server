Attribute VB_Name = "ModClimas"
Public IntensidadDeNubes As Byte
Public IntensidadDeLluvias As Byte
Public CapasLlueveEn As Integer
Public TimerMeteorologico As Byte
Public DuracionDeLLuvia As Integer
Public ServidorNublado As Boolean
Public ProbabilidadNublar As Byte
Public ProbabilidadLLuvia As Byte


Public Sub ResetMeteo()
Call AgregarAConsola("Servidor > Meteorologia reseteada")
frmMain.TimerMeteorologia.Enabled = True
frmMain.Truenos.Enabled = False
TimerMeteorologico = 30
ServidorNublado = False
Lloviendo = False

End Sub




Public Sub Nublar()

Dim ProbabilidadNubes As Long


'Empezar a nublar

'send nubes
'Intensidad variable

'iniciar timming de 1 a 3 minutos por si llueve
'enviar algun trueno



'Despues de 3 minutos

'probabilidad de lluvia
'enviar mega trueno y luz
'se larga a llover
'no llueve
'sacar nubes


End Sub
