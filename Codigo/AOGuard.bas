Attribute VB_Name = "AOGuard"
Option Explicit

Public AOG_STATUS As Byte

Private SMTP_HOST As String
Private SMTP_PORT As Integer
Private SMTP_AUTH As Byte
Private SMTP_SECURE As Byte
Private SMTP_USER As String
Private SMTP_PASS As String

Public Sub LoadAOGuardConfiguration()

    Dim ConfigFile As New clsIniManager
    Call ConfigFile.Initialize(IniPath & "AOGuard.ini")
        
    AOG_STATUS = val(ConfigFile.GetValue("INIT", "Enabled"))
        
    SMTP_HOST = ConfigFile.GetValue("INIT", "SMTP_HOST")
    SMTP_PORT = val(ConfigFile.GetValue("INIT", "SMTP_PORT"))
    SMTP_AUTH = val(ConfigFile.GetValue("INIT", "SMTP_AUTH"))
    SMTP_SECURE = val(ConfigFile.GetValue("INIT", "SMTP_AUTH"))
    SMTP_USER = ConfigFile.GetValue("INIT", "SMTP_USER")
    SMTP_PASS = ConfigFile.GetValue("INIT", "SMTP_PASS")
    
    Set ConfigFile = Nothing
    
End Sub

'------------------------------------------------------------------------------------------------
' Esto se va a encargar de chequear que el usuario se haya conectado desde una ubicacion segura.
'
' Se le dara acceso a la cuenta si:
'   - El HDSerial o la IP de la PC donde esta accediendo es igual a el que tenemos en la BD
'------------------------------------------------------------------------------------------------
Public Function VerificarOrigen(ByVal AccountID As Long, ByVal HD As Long, ByVal IP As String) As Boolean
    
    If Not IsNull(GetDBValue("account_guard", "code", "account_id", AccountID)) Then
        VerificarOrigen = False
        Exit Function
    End If
    
    Call MakeQuery("SELECT hd_serial, last_ip FROM account WHERE account_id = ?", False, AccountID)
    
    If QueryData Is Nothing Then
        VerificarOrigen = True
        Exit Function
    End If
    
    VerificarOrigen = (HD = QueryData!hd_serial Or IP = QueryData!last_ip)
    
    ' Mas adelante, si pinta ser mas exhaustivos podemos agregar chequeos de yokese...
    ' MAC, DNI, Numero de Tramite, lo que sea :)
    
End Function

'---------------------------------------------------------------------------------------------------
' Si VerificarOrigen = False, le notificamos al usuario que ponga el codigo que le mandamos al mail.
'---------------------------------------------------------------------------------------------------
Public Sub WriteGuardNotice(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.GuardNotice)
        Call .EndPacket
        
        Dim Codigo As String: Codigo = RandomString(5)
        
        ' Guardamos los valores en la base de datos
        Call MakeQuery("UPDATE account_guard SET code = ?, timestamp = ? WHERE account_id = ?", True, Codigo, Now(), UserList(UserIndex).AccountID)
        
        Debug.Print vbNewLine & "Codigo de Verificacion:" & Codigo & vbNewLine
        
        Call SendEmail(UserList(UserIndex).Cuenta, Codigo)
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub HandleGuardNoticeResponse(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Dim Codigo As String: Codigo = .incomingData.ReadASCIIString
        
        Dim DB_Codigo As String:    DB_Codigo = GetDBValue("account_guard", "code", "account_id", .AccountID)
        Dim DB_Timestamp As String: DB_Timestamp = GetDBValue("account_guard", "timestamp", "account_id", .AccountID)
        
        ' El codigo expira despues de 1 minuto.
        If DateDiff("s", Now(), DB_Timestamp) < 60 Then
        
            ' Le avisamos que expiro
            Call WriteErrorMsg(UserIndex, "El código de verificación ha expirado.")
            Call WriteDisconnect(UserIndex, True)
        
        Else ' El codigo NO expiro...
            
            ' Lo comparamos con lo que tenemos en la BD
            If Codigo = DB_Codigo Then
                Call WritePersonajesDeCuenta(UserIndex)
                Call WriteMostrarCuenta(UserIndex)
            
            Else
                Call WriteErrorMsg(UserIndex, "Codigo de verificación erroneo.")
                Call CloseSocket(UserIndex)
                
            End If
            
        End If

        ' Invalidamos el codigo
        Call MakeQuery("UPDATE account_guard SET code = ?, timestamp = ? WHERE account_id = ?", True, Null, Null, UserList(UserIndex).AccountID)
        
    End With
    
End Sub

' Source: https://accautomation.ca/how-to-send-email-to-smtp-server/
Sub SendEmail(ByVal Email As String, ByVal Codigo As String)

    On Error Resume Next
    
    If Not SMTP_HOST Or Not SMTP_PORT Or Not LenB(SMTP_USER) Or Not LenB(SMTP_PASS) Then Exit Sub
    
    Dim Schema As String
    
    Dim cdoMsg As Object
    Dim cdoConf As Object
    Dim cdoFields As Object
    
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    Set cdoFields = cdoConf.Fields
    
    ' Send one copy with Google SMTP server (with autentication)
    Schema = "http://schemas.microsoft.com/cdo/configuration/"
    
    cdoFields.Item(Schema & "sendusing") = 2
    cdoFields.Item(Schema & "smtpserver") = SMTP_HOST
    cdoFields.Item(Schema & "smtpserverport") = SMTP_PORT
    cdoFields.Item(Schema & "smtpauthenticate") = SMTP_AUTH
    cdoFields.Item(Schema & "sendusername") = SMTP_USER
    cdoFields.Item(Schema & "sendpassword") = SMTP_PASS
    cdoFields.Item(Schema & "smtpusessl") = SMTP_SECURE
    
    Call cdoFields.Update

    With cdoMsg
    
        .To = Email
        .From = "argentum20@ao20.com.ar"
        .Subject = "Argentum Guard - Acceso desde un nuevo dispositivo"
        
        ' Body of message can be any HTML code
        .HTMLBody = "Codigo: " & Codigo
        
        Set .Configuration = cdoConf
        
        ' Send the message
        Call .Send
    End With

    'Check for errors and display message
    If Err.Number <> 0 Then
        Call RegistrarError(500, "Error al enviar correo a " & Email & vbNewLine & Err.Description, "AOGuard.SendMail")
    End If

    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set cdoFields = Nothing

End Sub
