Attribute VB_Name = "Logging"
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



Public Sub LogearEventoDeSubasta(Logeo As String)
        
        On Error GoTo LogearEventoDeSubasta_Err
        

        Dim n As Integer

100     n = FreeFile
102     Open App.Path & "\LOGS\subastas.log" For Append Shared As n
104     Print #n, Logeo
106     Close #n

        
        Exit Sub

LogearEventoDeSubasta_Err:
108     Call TraceError(Err.Number, Err.Description, "ModSubasta.LogearEventoDeSubasta", Erl)

        
End Sub

Sub LogBan(ByVal BannedIndex As Integer, ByVal userindex As Integer, ByVal Motivo As String)
        
        On Error GoTo LogBan_Err
        

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(userindex).Name)
102     Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)

        'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
        Dim mifile As Integer

104     mifile = FreeFile
106     Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
108     Print #mifile, UserList(BannedIndex).Name
110     Close #mifile

        
        Exit Sub

LogBan_Err:
112     Call TraceError(Err.Number, Err.Description, "ES.LogBan", Erl)

        
End Sub


Public Sub LogCreditosPatreon(Desc As String)
        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile
    
102     Open App.Path & "\logs\Monetization\CreditosPatreon.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " - " & Desc
106     Close #nfile
     
        Exit Sub
    
ErrHandler:

End Sub

Public Sub LogShopTransactions(Desc As String)
        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile
    
102     Open App.Path & "\logs\Monetization\Shop\Transactions.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " - " & Desc
106     Close #nfile
                 
        Exit Sub
    
ErrHandler:

End Sub

Public Sub LogShopErrors(Desc As String)
        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile
    
102     Open App.Path & "\logs\Monetization\Shop\Errors.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " - " & Desc
106     Close #nfile
                 
        Exit Sub
    
ErrHandler:

End Sub


Public Sub LogEdicionPaquete(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\EdicionPaquete.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

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

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
104     Print #nfile, Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
104     Print #nfile, Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogError(Desc As String)

100     Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    
102     Open App.Path & "\logs\errores.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

End Sub

Public Sub LogPerformance(Desc As String)

100     Dim nfile As Integer: nfile = FreeFile ' obtenemos un canal
    
102     Open App.Path & "\logs\Performance.log" For Append Shared As #nfile
104         Print #nfile, Date & " " & Time & " " & Desc
106     Close #nfile

        Exit Sub

End Sub

Public Sub LogConsulta(Desc As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\ConsultasGM.log" For Append Shared As #nfile
104     Print #nfile, Date & " - " & Time & " - " & Desc
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogClanes(ByVal str As String)
        
        On Error GoTo LogClanes_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogClanes_Err:
108     Call TraceError(Err.Number, Err.Description, "General.LogClanes", Erl)

        
End Sub


Public Sub LogDesarrollo(ByVal str As String)
        
        On Error GoTo LogDesarrollo_Err
        

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        
        Exit Sub

LogDesarrollo_Err:
108     Call TraceError(Err.Number, Err.Description, "General.LogDesarrollo", Erl)

        
End Sub

Public Sub LogGM(nombre As String, texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
        'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
102     Open App.Path & "\logs\" & nombre & ".log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogPremios(GM As String, UserName As String, ByVal ObjIndex As Integer, ByVal Cantidad As Integer, Motivo As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal

102     Open App.Path & "\logs\PremiosOtorgados.log" For Append Shared As #nfile
104     Print #nfile, "[" & GM & "]" & vbNewLine
106     Print #nfile, Date & " " & Time & vbNewLine
108     Print #nfile, "Item: " & ObjData(ObjIndex).Name & " (" & ObjIndex & ") Cantidad: " & Cantidad & vbNewLine
110     Print #nfile, "Motivo: " & Motivo & vbNewLine & vbNewLine
112     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogDatabaseError(Desc As String)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
    
102     Open App.Path & "\logs\Database.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " - " & Desc
106     Close #nfile
     
108     Debug.Print "Error en la BD: " & Desc & vbNewLine & _
            "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
            
        Exit Sub
    
ErrHandler:

End Sub

Public Sub LogHackAttemp(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
104     Print #nfile, "----------------------------------------------------------"
106     Print #nfile, Date & " " & Time & " " & texto
108     Print #nfile, "----------------------------------------------------------"
110     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogCheating(texto As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\CH.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & texto
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

