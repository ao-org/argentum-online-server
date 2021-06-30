Attribute VB_Name = "AOGuard"
Option Explicit

Public AOG_STATUS As Byte
Private AOG_EXPIRE As Byte
Private AOG_RESEND As Long

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
    AOG_EXPIRE = val(ConfigFile.GetValue("INIT", "CodeExpiresInSeconds"))
    
    AOG_RESEND = val(ConfigFile.GetValue("INIT", "CodeResendInterval")) * 10000
    If AOG_RESEND = 0 Then AOG_RESEND = 50000
    
    SMTP_HOST = ConfigFile.GetValue("SMTP", "HOST")
    SMTP_PORT = val(ConfigFile.GetValue("SMTP", "PORT"))
    SMTP_AUTH = val(ConfigFile.GetValue("SMTP", "AUTH"))
    SMTP_SECURE = val(ConfigFile.GetValue("SMTP", "SECURE"))
    SMTP_USER = ConfigFile.GetValue("SMTP", "USER")
    SMTP_PASS = ConfigFile.GetValue("SMTP", "PASS")
    
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

Private Sub GenerarCodigo(ByVal UserIndex As Integer)
    
    Dim Codigo As String
    Dim NuevoCodigo As Boolean
    
    With UserList(UserIndex)
        
        Call MakeQuery("SELECT * FROM account_guard WHERE account_id = ?", False, .AccountID)
     
        ' NO tiene codigo
        If QueryData Is Nothing Then

            NuevoCodigo = True
        
        ' Tiene codigo, pero ya expiro...
        ElseIf AOG_EXPIRE <> 0 And DateDiff("s", Now(), QueryData!TimeStamp) > AOG_EXPIRE Then
            
            NuevoCodigo = True
            
        Else ' Si ya tiene codigo y NO expiro...
            
            NuevoCodigo = False

        End If
        
        If NuevoCodigo Then
        
            ' Generamos un nuevo codigo
            Codigo = RandomString(5)
            
            
            ' Lo guardamos en la BD
            Call MakeQuery("REPLACE INTO account_guard (account_id, code) VALUES (?, ?)", True, .AccountID, Codigo)
        
        Else
            
            ' Usamos el codigo vigente
            Codigo = QueryData!code
            
        End If
        
        Debug.Print "Codigo de Verificacion: " & Codigo & vbNewLine
        
        ' Enviamos el mail con el codigo
        Call SendEmail(.Cuenta, Codigo, .IP)
        
    End With
    
End Sub

'---------------------------------------------------------------------------------------------------
' Si VerificarOrigen = False, le notificamos al usuario que ponga el codigo que le mandamos al mail.
'---------------------------------------------------------------------------------------------------
Public Sub WriteGuardNotice(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.GuardNotice)
        Call .EndPacket
        
        Call GenerarCodigo(UserIndex)
    
    End With
    
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub HandleGuardNoticeResponse(ByVal UserIndex As Integer)
    
    On Error GoTo HandleGuardNoticeResponse_Err:
    
    With UserList(UserIndex)
        
        Dim Exito As Boolean
        
        Dim Codigo As String: Codigo = .incomingData.ReadASCIIString
        
        Dim DB_Codigo As String:    DB_Codigo = GetDBValue("account_guard", "code", "account_id", .AccountID)
        Dim DB_Timestamp As String: DB_Timestamp = GetDBValue("account_guard", "timestamp", "account_id", .AccountID)
        
        ' El codigo expira despues de 1 minuto.
        If AOG_EXPIRE <> 0 And DateDiff("s", Now(), DB_Timestamp) < AOG_EXPIRE Then
        
            ' Le avisamos que expiro
            Call WriteShowMessageBox(UserIndex, "El código de verificación ha expirado.")
            
            ' Invalidamos el codigo
            Call MakeQuery("DELETE FROM account_guard WHERE account_id = ?", True, UserList(UserIndex).AccountID)
            
            ' Lo kickeamos.
             Call CloseSocket(UserIndex)
            
        Else ' El codigo NO expiro...
            
            ' Lo comparamos con lo que tenemos en la BD
            If Codigo = DB_Codigo Then
                Call WritePersonajesDeCuenta(UserIndex)
                Call WriteMostrarCuenta(UserIndex)
                
                ' Invalidamos el codigo
                Call MakeQuery("DELETE FROM account_guard WHERE account_id = ?", True, UserList(UserIndex).AccountID)
                
            Else
                ' Le avisamos
                Call WriteShowMessageBox(UserIndex, "El código de verificación ha expirado.")
                
                ' Lo kickeamos.
                Call CloseSocket(UserIndex)
            End If
            
        End If
 
    End With
    
    Exit Sub

HandleGuardNoticeResponse_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuardNoticeResponse", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

' Source: https://accautomation.ca/how-to-send-email-to-smtp-server/
Sub SendEmail(ByVal Email As String, ByVal Codigo As String, ByVal IP As String)

    On Error Resume Next
    
    If LenB(SMTP_HOST) = 0 Or _
        LenB(SMTP_PORT) = 0 Or _
        LenB(SMTP_USER) = 0 Or _
        LenB(SMTP_PASS) = 0 Then Exit Sub
    
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
        .From = "guardian@ao20.com.ar"
        .Subject = "Argentum Guard - Acceso desde un nuevo dispositivo"
        
        ' Body of message can be any HTML code
        .HTMLBody = "Hemos detectado un intento de acceso a tu cuenta desde un dispositivo desconocido <br /><br />" & _
                    "IP: " & IP & "<br /><br />" & _
                    "Si fuiste tu, te aparecerá un dialogo donde tendrás que ingresar el siguiente código: <strong>" & Codigo & "</strong>" & "<br />" & _
                    "<strong>Si NO fuiste tu, ignora este mensaje y considera cambiar tu contraseña</strong>" & "<br /><br />" & _
                    "Atentamente, el Staff de Argentum20"
                    
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
