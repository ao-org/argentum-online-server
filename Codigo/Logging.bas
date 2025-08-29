Attribute VB_Name = "Logging"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
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
 
 Private Enum type_log
    e_LogearEventoDeSubasta = 0
    e_LogBan = 1
    e_LogCreditosPatreon = 2
    e_LogShopTransactions = 3
    e_LogShopErrors = 4
    e_LogEdicionPaquete = 5
    e_LogMacroServidor = 6
    e_LogMacroCliente = 7
    e_LogVentaCasa = 8
    e_LogCriticEvent = 9
    e_LogEjercitoReal = 10
    e_LogEjercitoCaos = 11
    e_LogError = 12
    e_LogPerformance = 13
    e_LogConsulta = 14
    e_LogClanes = 15
    e_LogGM = 16
    e_LogPremios = 17
    e_LogDatabaseError = 18
    e_LogSecurity = 19
 End Enum
Private Type t_CircularBuffer
    currentIndex As Integer
    Messages() As String
    Size As Integer
End Type
Public CircularLogBuffer As t_CircularBuffer

Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" ( _
 ByVal lpUNCServerName As String, _
 ByVal lpSourceName As String) As Long


Public Sub InitializeCircularLogBuffer(Optional ByVal Size As Integer = 30)
    On Error Goto InitializeCircularLogBuffer_Err
    CircularLogBuffer.Size = Size
    CircularLogBuffer.currentIndex = 0
    ReDim CircularLogBuffer.Messages(0 To Size)
    Exit Sub
InitializeCircularLogBuffer_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.InitializeCircularLogBuffer", Erl)
End Sub

Public Sub AddLogToCircularBuffer(Message As String)
    On Error Goto AddLogToCircularBuffer_Err
    CircularLogBuffer.currentIndex = CircularLogBuffer.currentIndex + 1
    CircularLogBuffer.currentIndex = (CircularLogBuffer.currentIndex Mod CircularLogBuffer.Size)
    CircularLogBuffer.Messages(CircularLogBuffer.currentIndex) = Message
    Exit Sub
AddLogToCircularBuffer_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.AddLogToCircularBuffer", Erl)
End Sub

Public Function GetLastMessages() As String()
    On Error Goto GetLastMessages_Err
    Dim errorList() As String
    ReDim errorList(CircularLogBuffer.Size)
    Dim i As Integer
    Dim circularIndex As Integer
    For i = 1 To CircularLogBuffer.Size
        circularIndex = ((CircularLogBuffer.currentIndex + i) Mod CircularLogBuffer.Size)
        errorList(i) = CircularLogBuffer.Messages(circularIndex)
    Next i
    GetLastMessages = errorList
    Exit Function
GetLastMessages_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.GetLastMessages", Erl)
End Function



Public Sub LogThis(nErrNo As Long, sLogMsg As String, eventType As LogEventTypeConstants)
    On Error Goto LogThis_Err
    Dim hEvent As Long
    hEvent = RegisterEventSource("", "Argentum20")
    If eventType = vbLogEventTypeWarning Or eventType = vbLogEventTypeError Then
        Call AddLogToCircularBuffer(sLogMsg)
    End If
    Call ReportEvent(hEvent, eventType, 0, 20, 0, 1, Len(sLogMsg), nErrNo & " - " & sLogMsg, 0)
    Exit Sub
LogThis_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogThis", Erl)
End Sub

Public Sub LogearEventoDeSubasta(s As String)
    On Error Goto LogearEventoDeSubasta_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogearEventoDeSubasta, "[Subastas.log] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogearEventoDeSubasta_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogearEventoDeSubasta", Erl)
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
    On Error Goto LogBan_Err
On Error GoTo ErrHandler
        Dim s As String
        s = UserList(BannedIndex).name & " BannedBy " & UserList(UserIndex).name & " Reason " & Motivo
        Call LogThis(type_log.e_LogBan, "[Bans] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogBan_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogBan", Erl)
End Sub


Public Sub LogCreditosPatreon(Desc As String)
    On Error Goto LogCreditosPatreon_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogCreditosPatreon, "[MonetizationCreditosPatreon.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogCreditosPatreon_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogCreditosPatreon", Erl)
End Sub

Public Sub LogShopTransactions(Desc As String)
    On Error Goto LogShopTransactions_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogShopTransactions, "[MonetizationShopTransactions.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogShopTransactions_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogShopTransactions", Erl)
End Sub

Public Sub LogShopErrors(Desc As String)
    On Error Goto LogShopErrors_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogShopErrors, "[MonetizationShopErrors.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
    Exit Sub
LogShopErrors_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogShopErrors", Erl)
End Sub


Public Sub LogEdicionPaquete(texto As String)
    On Error Goto LogEdicionPaquete_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEdicionPaquete, "[EdicionPaquete.log] " & texto, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
    Exit Sub
LogEdicionPaquete_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogEdicionPaquete", Erl)
End Sub

Public Sub LogMacroServidor(texto As String)
    On Error Goto LogMacroServidor_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogMacroServidor, "[MacroServidor] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogMacroServidor_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogMacroServidor", Erl)
End Sub

Public Sub LogMacroCliente(texto As String)
    On Error Goto LogMacroCliente_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogMacroCliente, "[MacroCliente] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogMacroCliente_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogMacroCliente", Erl)
End Sub
Public Sub logVentaCasa(ByVal texto As String)
    On Error Goto logVentaCasa_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogVentaCasa, "[Propiedades] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
logVentaCasa_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.logVentaCasa", Erl)
End Sub


Public Sub LogCriticEvent(Desc As String)
    On Error Goto LogCriticEvent_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogCriticEvent, "[Eventos.log] " & Desc, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
    Exit Sub
LogCriticEvent_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogCriticEvent", Erl)
End Sub

Public Sub LogEjercitoReal(Desc As String)
    On Error Goto LogEjercitoReal_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEjercitoReal, "[EjercitoReal.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogEjercitoReal_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogEjercitoReal", Erl)
End Sub

Public Sub LogEjercitoCaos(Desc As String)
    On Error Goto LogEjercitoCaos_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEjercitoCaos, "[EjercitoCaos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogEjercitoCaos_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogEjercitoCaos", Erl)
End Sub

Public Sub LogError(Desc As String)
    On Error Goto LogError_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogError, "[Errores.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
    Exit Sub
LogError_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogError", Erl)
End Sub

Public Sub LogPerformance(Desc As String)
    On Error Goto LogPerformance_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogPerformance, "[Performance.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogPerformance_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogPerformance", Erl)
End Sub

Public Sub LogConsulta(Desc As String)
    On Error Goto LogConsulta_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogConsulta, "[obtenemos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogConsulta_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogConsulta", Erl)
End Sub

Public Sub LogClanes(ByVal str As String)
    On Error Goto LogClanes_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogClanes, "[Clans.log] " & str, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogClanes_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogClanes", Erl)
End Sub
Public Sub LogGM(name As String, Desc As String)
    On Error Goto LogGM_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogGM, "[" & name & "] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogGM_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogGM", Erl)
End Sub

Public Sub LogPremios(GM As String, username As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)
    On Error Goto LogPremios_Err
On Error GoTo ErrHandler
        Dim s As String
        s = "Item: " & ObjData(ObjIndex).name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine _
        & "Motivo: " & Motivo & vbNewLine & vbNewLine
        Call LogThis(type_log.e_LogPremios, s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
    Exit Sub
LogPremios_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogPremios", Erl)
End Sub

Public Sub LogDatabaseError(Desc As String)
    On Error Goto LogDatabaseError_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogDatabaseError, "[Database.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
    Exit Sub
LogDatabaseError_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogDatabaseError", Erl)
End Sub

Public Sub LogSecurity(str As String)
    On Error Goto LogSecurity_Err
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogSecurity, "[Cheating.log] " & str, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
    Exit Sub
LogSecurity_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.LogSecurity", Erl)
End Sub

Public Sub TraceError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    On Error Goto TraceError_Err
    #If DEBUGGING = 1 Then
        Debug.Print "TraceError: " & Descripcion & " " & Componente
    #End If
    Call LogThis(Numero, "Error number: " & Numero & " | Description: " & Descripcion & vbNewLine & "Component: " & Componente & " | Line number: " & Linea, vbLogEventTypeError)
    Exit Sub
TraceError_Err:
    Call TraceError(Err.Number, Err.Description, "Logging.TraceError", Erl)
End Sub
