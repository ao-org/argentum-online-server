Attribute VB_Name = "Logging"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
'    Copyright (C) 2002 Mï¿½rquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Private Declare Function ReportEvent _
                Lib "advapi32.dll" _
                Alias "ReportEventA" (ByVal hEventLog As Long, _
                                      ByVal wType As Integer, _
                                      ByVal wCategory As Integer, _
                                      ByVal dwEventID As Long, _
                                      ByVal lpUserSid As Long, _
                                      ByVal wNumStrings As Integer, _
                                      ByVal dwDataSize As Long, _
                                      plpStrings As String, _
                                      lpRawData As Long) As Long

Private Enum eType_Log
    EventoDeSubasta = 0
    Ban = 1
    CreditosPatreon = 2
    ShopTransactions = 3
    ShopErrors = 4
    EdicionPaquete = 5
    MacroServidor = 6
    MacroCliente = 7
    VentaCasa = 8
    CriticEvent = 9
    EjercitoReal = 10
    EjercitoCaos = 11
    Error = 12
    Performance = 13
    Consulta = 14
    Clanes = 15
    GM = 16
    Premios = 17
    DatabaseError = 18
    Security = 19
    BankTransfer = 20
End Enum

Private Type t_CircularBuffer
    currentIndex As Integer
    Messages() As String
    size As Integer
End Type

Public CircularLogBuffer As t_CircularBuffer
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long

Private m_hEventLog As Long

Public Sub InitializeCircularLogBuffer(Optional ByVal size As Integer = 30)
    CircularLogBuffer.size = size
    CircularLogBuffer.currentIndex = 0
    ReDim CircularLogBuffer.Messages(0 To size)
End Sub

Public Sub InitializeLogging()
    m_hEventLog = RegisterEventSource("", "Argentum20")
    Call InitializeCircularLogBuffer(30)
End Sub

Public Sub ShutdownLogging()
    If m_hEventLog <> 0 Then
        DeregisterEventSource m_hEventLog
        m_hEventLog = 0
    End If
End Sub

Public Sub AddLogToCircularBuffer(Message As String)
    CircularLogBuffer.currentIndex = CircularLogBuffer.currentIndex + 1
    CircularLogBuffer.currentIndex = (CircularLogBuffer.currentIndex Mod CircularLogBuffer.size)
    CircularLogBuffer.Messages(CircularLogBuffer.currentIndex) = Message
End Sub

Public Function GetLastMessages() As String()
    Dim errorList() As String
    ReDim errorList(CircularLogBuffer.size)
    Dim i             As Integer
    Dim circularIndex As Integer
    For i = 1 To CircularLogBuffer.size
        circularIndex = ((CircularLogBuffer.currentIndex + i) Mod CircularLogBuffer.size)
        errorList(i) = CircularLogBuffer.Messages(circularIndex)
    Next i
    GetLastMessages = errorList
End Function

Public Sub LogThis(nErrNo As Long, sLogMsg As String, eventType As LogEventTypeConstants)
    If m_hEventLog = 0 Then Call InitializeLogging
    If eventType = vbLogEventTypeWarning Or eventType = vbLogEventTypeError Then
        Call AddLogToCircularBuffer(sLogMsg)
    End If
    Call ReportEvent(m_hEventLog, eventType, 0, nErrNo, 0, 1, Len(sLogMsg), sLogMsg, 0)
End Sub

Public Sub LogearEventoDeSubasta(s As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EventoDeSubasta, "[Subastas] " & s, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
    On Error GoTo ErrHandler
    Dim s As String
    s = UserList(BannedIndex).name & " BannedBy " & UserList(UserIndex).name & " Reason " & Motivo
    Call LogThis(eType_Log.Ban, "[Bans] " & s, vbLogEventTypeWarning)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogCreditosPatreon(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.CreditosPatreon, "[CreditosPatreon] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogShopTransactions(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.ShopTransactions, "[ShopTransactions] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogShopErrors(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.ShopErrors, "[ShopErrors] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogEdicionPaquete(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EdicionPaquete, "[EdicionPaquete] " & texto, vbLogEventTypeWarning)
    Exit Sub
ErrHandler:
End Sub
Public Sub LogInfoServidor(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.MacroServidor, "[Info] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroServidor(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.MacroServidor, "[MacroServidor] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogMacroCliente(texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.MacroCliente, "[MacroCliente] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub logVentaCasa(ByVal texto As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.VentaCasa, "[Propiedades] " & texto, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogCriticEvent(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.CriticEvent, "[CriticEvent] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoReal(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EjercitoReal, "[EjercitoReal] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogEjercitoCaos(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.EjercitoCaos, "[EjercitoCaos] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogError(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Error, "[Error] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogPerformance(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Performance, "[Performance] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogConsulta(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Consulta, "[Consulta] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogClanes(ByVal str As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Clanes, "[Clanes] " & str, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogGM(name As String, Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.GM, "[" & name & "] " & Desc, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogPremios(GM As String, username As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)
    On Error GoTo ErrHandler
    Dim s As String
    s = "Item: " & ObjData(ObjIndex).name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine & "Motivo: " & Motivo & vbNewLine & vbNewLine
    Call LogThis(eType_Log.Premios, s, vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogDatabaseError(Desc As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.DatabaseError, "[Database] " & Desc, vbLogEventTypeError)
    Exit Sub
ErrHandler:
End Sub

Public Sub LogSecurity(str As String)
    On Error GoTo ErrHandler
    Call LogThis(eType_Log.Security, "[Security] " & str, vbLogEventTypeWarning)
    Exit Sub
ErrHandler:
End Sub
Public Sub LogBankTransfer(ByVal originUser As String, ByVal targetUser As String, ByVal amount As Long, Optional ByVal receiverOnline As Boolean = False)
    On Error GoTo ErrHandler
    Dim transferContext As String
    If receiverOnline Then
        transferContext = "destinatario en lï¿½nea"
    Else
        transferContext = "destinatario fuera de lï¿½nea"
    End If
    Call LogThis(eType_Log.BankTransfer, "[BankTransfer] " & originUser & " transfiriï¿½ " & amount & " monedas a " & targetUser & " (" & transferContext & ")", vbLogEventTypeInformation)
    Exit Sub
ErrHandler:
End Sub



Public Sub TraceError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    #If DEBUGGING = 1 Then
        Debug.Print "TraceError: " & Descripcion & " " & Componente
    #End If
    Call LogThis(Numero, "Error number: " & Numero & " | Description: " & Descripcion & vbNewLine & "Component: " & Componente & " | Line number: " & Linea, vbLogEventTypeError)
End Sub
