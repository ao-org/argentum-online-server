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
    CircularLogBuffer.Size = Size
    CircularLogBuffer.currentIndex = 0
    ReDim CircularLogBuffer.Messages(0 To Size)
End Sub

Public Sub AddLogToCircularBuffer(Message As String)
    CircularLogBuffer.currentIndex = CircularLogBuffer.currentIndex + 1
    CircularLogBuffer.currentIndex = (CircularLogBuffer.currentIndex Mod CircularLogBuffer.Size)
    CircularLogBuffer.Messages(CircularLogBuffer.currentIndex) = Message
End Sub

Public Function GetLastMessages() As String()
    Dim errorList() As String
    ReDim errorList(CircularLogBuffer.Size)
    Dim i As Integer
    Dim circularIndex As Integer
    For i = 1 To CircularLogBuffer.Size
        circularIndex = ((CircularLogBuffer.currentIndex + i) Mod CircularLogBuffer.Size)
        errorList(i) = CircularLogBuffer.Messages(circularIndex)
    Next i
    GetLastMessages = errorList
End Function



Public Sub LogThis(nErrNo As Long, sLogMsg As String, eventType As LogEventTypeConstants)
    Dim hEvent As Long
    hEvent = RegisterEventSource("", "Argentum20")
    If eventType = vbLogEventTypeWarning Or eventType = vbLogEventTypeError Then
        Call AddLogToCircularBuffer(sLogMsg)
    End If
    Call ReportEvent(hEvent, eventType, 0, 20, 0, 1, Len(sLogMsg), nErrNo & " - " & sLogMsg, 0)
End Sub

Public Sub LogearEventoDeSubasta(s As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogearEventoDeSubasta, "[Subastas.log] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
On Error GoTo ErrHandler
        Dim s As String
        s = UserList(BannedIndex).name & " BannedBy " & UserList(UserIndex).name & " Reason " & Motivo
        Call LogThis(type_log.e_LogBan, "[Bans] " & s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogCreditosPatreon(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogCreditosPatreon, "[MonetizationCreditosPatreon.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogShopTransactions(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogShopTransactions, "[MonetizationShopTransactions.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogShopErrors(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogShopErrors, "[MonetizationShopErrors.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogEdicionPaquete(texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEdicionPaquete, "[EdicionPaquete.log] " & texto, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroServidor(texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogMacroServidor, "[MacroServidor] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroCliente(texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogMacroCliente, "[MacroCliente] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogVentaCasa, "[Propiedades] " & texto, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogCriticEvent, "[Eventos.log] " & Desc, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEjercitoReal, "[EjercitoReal.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogEjercitoCaos, "[EjercitoCaos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogError(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogError, "[Errores.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogPerformance(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogPerformance, "[Performance.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogConsulta(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogConsulta, "[obtenemos.log] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogClanes(ByVal str As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogClanes, "[Clans.log] " & str, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub
Public Sub LogGM(name As String, Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogGM, "[" & name & "] " & Desc, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogPremios(GM As String, username As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)
On Error GoTo ErrHandler
        Dim s As String
        s = "Item: " & ObjData(ObjIndex).name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine _
        & "Motivo: " & Motivo & vbNewLine & vbNewLine
        Call LogThis(type_log.e_LogPremios, s, vbLogEventTypeInformation)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogDatabaseError(Desc As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogDatabaseError, "[Database.log] " & Desc, vbLogEventTypeError)
        Exit Sub
ErrHandler:
End Sub

Public Sub LogSecurity(str As String)
On Error GoTo ErrHandler
        Call LogThis(type_log.e_LogSecurity, "[Cheating.log] " & str, vbLogEventTypeWarning)
        Exit Sub
ErrHandler:
End Sub

Public Sub TraceError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    'Start append text to file
    Call LogThis(Numero, "Error number: " & Numero & " | Description: " & Descripcion & vbNewLine & "Component: " & Componente & " | Line number: " & Linea, vbLogEventTypeError)
End Sub
