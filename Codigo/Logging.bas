Attribute VB_Name = "Logging"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit


Private Declare Function ReportEvent _
 Lib "advapi32.dll" Alias "ReportEventA" ( _
 ByVal hEventLog As Long, _
 ByVal wType As Integer, _
 ByVal wCategory As Integer, _
 ByVal dwEventID As Long, _
 ByVal lpUserSid As Long, _
 ByVal wNumStrings As Integer, _
 ByVal dwDataSize As Long, _
 plpStrings As String, _
 lpRawData As Long) As Long

Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" ( _
 ByVal lpUNCServerName As String, _
 ByVal lpSourceName As String) As Long

Public Sub LogThis(nErrNo As Long, sLogMsg As String, EventType As LogEventTypeConstants)
    Dim hEvent As Long
    hEvent = RegisterEventSource("", "Argentum20")
    Call ReportEvent(hEvent, EventType, 0, nErrNo, 0, 1, Len(sLogMsg), sLogMsg, 0)
End Sub

Public Sub LogearEventoDeSubasta(s As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Subastas.log] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal userindex As Integer, ByVal Motivo As String)
On Error GoTo ErrHandler
        Dim s As String
        s = UserList(BannedIndex).name & " BannedBy " & UserList(UserIndex).name & " Reason " & Motivo
        Call LogThis(0, "[Bans] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogCreditosPatreon(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[MonetizationCreditosPatreon.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogShopTransactions(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[MonetizationShopTransactions.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogShopErrors(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[MonetizationShopErrors.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogEdicionPaquete(texto As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[EdicionPaquete.log] " & texto, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroServidor(texto As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[MacroServidor] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroCliente(texto As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[MacroCliente] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Propiedades] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Eventos.log] " & Desc, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[EjercitoReal.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[EjercitoCaos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogError(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Errores.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogPerformance(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Performance.log] " & desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogConsulta(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[obtenemos.log] " & desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogClanes(ByVal str As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Clans.log] " & str, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogGM(name As String, desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[" & name & "] " & desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogPremios(GM As String, UserName As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)
On Error GoTo ErrHandler
        Dim s As String
        s = "Item: " & ObjData(ObjIndex).name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine _
        & "Motivo: " & Motivo & vbNewLine & vbNewLine
        Call LogThis(0, s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogDatabaseError(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Database.log] " & desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogSecurity(str As String)
On Error GoTo ErrHandler
        Call LogThis(0, "[Cheating.log] " & str, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub TraceError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
On Error GoTo ErrHandler
 Dim s As String
 s = "Response Code: " & Numero & " Response Error Description: " & Descripcion & " Response Contents: " & Componente
 Call LogThis(0, "[Trace.log] " & s, vbLogEventTypeError)
ErrHandler:
End Sub
