Attribute VB_Name = "AOGuard"
Option Explicit

Private Const MAX_CODE_RESEND_COUNT As Byte = 10

' Configuracion - Argentum Guard
Public AOG_STATUS                   As Byte
Private AOG_EXPIRE                  As Long
Private AOG_RESEND_INTERVAL         As Long
Private TRANSPORT_METHOD            As String

' Configuracion API
Private API_ENDPOINT                As String
Private API_KEY                     As String

' Configuracion SMTP interno
' (no me gusta xq bloquea el hilo pero bueh, puede servir para salir del paso)
Private SMTP_HOST                   As String
Private SMTP_PORT                   As Integer
Private SMTP_AUTH                   As Byte
Private SMTP_SECURE                 As Byte
Private SMTP_USER                   As String
Private SMTP_PASS                   As String

Public Sub LoadAOGuardConfiguration()
    
    If Not FileExist(IniPath & "AOGuard.ini", vbNormal) Then
        AOG_STATUS = 0
        Exit Sub

    End If
    
    Dim ConfigFile As New clsIniManager
    Call ConfigFile.Initialize(IniPath & "AOGuard.ini")
        
    AOG_STATUS = val(ConfigFile.GetValue("INIT", "Enabled"))
    AOG_EXPIRE = val(ConfigFile.GetValue("INIT", "CodeExpiresInSeconds"))
    
    AOG_RESEND_INTERVAL = val(ConfigFile.GetValue("INIT", "CodeResendInterval"))

    If AOG_RESEND_INTERVAL = 0 Then AOG_RESEND_INTERVAL = 50000
    
    TRANSPORT_METHOD = UCase$(ConfigFile.GetValue("INIT", "TransportMethod"))

    Select Case TRANSPORT_METHOD
    
        Case "API"
            ' Configuracion API
            API_ENDPOINT = ConfigFile.GetValue("API", "Endpoint")
            API_KEY = ConfigFile.GetValue("API", "Key")
            
        Case "SMTP"
            ' Configuracion SMTP interno
            ' (no me gusta xq bloquea el hilo pero bueh, puede servir para salir del paso)
            SMTP_HOST = ConfigFile.GetValue("SMTP", "HOST")
            SMTP_PORT = val(ConfigFile.GetValue("SMTP", "PORT"))
            SMTP_AUTH = val(ConfigFile.GetValue("SMTP", "AUTH"))
            SMTP_SECURE = val(ConfigFile.GetValue("SMTP", "SECURE"))
            SMTP_USER = ConfigFile.GetValue("SMTP", "USER")
            SMTP_PASS = ConfigFile.GetValue("SMTP", "PASS")
            
    End Select
    
    Set ConfigFile = Nothing
    
End Sub

'------------------------------------------------------------------------------------------------
' Esto se va a encargar de chequear que el usuario se haya conectado desde una ubicacion segura.
'
' Se le dara acceso a la cuenta si:
'   - El HDSerial o la IP de la PC donde esta accediendo es igual a el que tenemos en la BD
'------------------------------------------------------------------------------------------------
Public Function VerificarOrigen(ByVal AccountID As Long, ByVal HD As Long, ByVal IP As String) As Boolean

        On Error GoTo VerificarOrigen_Err
    
100     If LenB(GetDBValue("account_guard", "code", "account_id", AccountID)) <> 0 Then
102         VerificarOrigen = False
            Exit Function

        End If
    
104     Call MakeQuery("SELECT hd_serial, last_ip FROM account WHERE id = ?", False, AccountID)
    
106     If QueryData Is Nothing Then
108         VerificarOrigen = True
            Exit Function

        End If
    
110     VerificarOrigen = (HD = QueryData!hd_serial Or IP = QueryData!last_ip)
    
        ' Mas adelante, si pinta ser mas exhaustivos podemos agregar chequeos de yokese...
        ' MAC, DNI, Numero de Tramite, lo que sea :)
    
        Exit Function

VerificarOrigen_Err:
        Call RegistrarError(Err.Number, Err.Description, "Protocol.VerificarOrigen", Erl)

End Function

'---------------------------------------------------------------------------------------------------
' Le enviamos el codigo de verificacion al usuario si la situacion lo requiere
'---------------------------------------------------------------------------------------------------
Private Sub EnviarCodigo(ByVal UserIndex As Integer)

    On Error GoTo EnviarCodigo_Err
    
    Dim EnviarCode As Boolean
    
    With UserList(UserIndex)
            
        Call MakeQuery("SELECT TIMESTAMPDIFF(SECOND, `code_last_sent`, CURRENT_TIMESTAMP) AS delta_time, code_resend_attempts FROM account_guard WHERE account_id = ?", False, .AccountID)
        
        ' Hay registros en `account_guard` = tiene codigo = me fijo si le mando o no
        If Not QueryData Is Nothing Then
            
            ' Ya te dije X veces que esperes un toque! Si no lo haces, sos alto bot!
            If val(QueryData!code_resend_attempts) > MAX_CODE_RESEND_COUNT Then
                EnviarCode = False
                Call CloseSocket(UserIndex)
                Exit Sub
                    
            End If
                
            ' Establecemos un intervalo de tiempo para volver a mandarle el codigo al usuario
            If QueryData!delta_time > AOG_RESEND_INTERVAL Then
                    
                EnviarCode = True
                    
                Call MakeQuery("UPDATE account_guard SET code_last_sent = CURRENT_TIMESTAMP, code_resend_attempts = 0 WHERE account_id = ?", True, .AccountID)
                        
                Call WriteShowMessageBox(UserIndex, "Te hemos enviado un correo con el código de verificacion a tu correo. " & _
                                                    "Si no lo encuentras, revisa la carpeta de SPAM. " & _
                                                    "Si no te ha llegado, intenta nuevamente en " & AOG_RESEND_INTERVAL \ 10000 & " segundos")
                        
            Else
                
                EnviarCode = False
                
                Call MakeQuery("UPDATE account_guard SET code_resend_attempts = code_resend_attempts + 1 WHERE account_id = ?", True, .AccountID)
                    
                Call WriteShowMessageBox(UserIndex, "Ya te hemos enviado un correo con el código de verificacion. " & _
                                                    "Si no te ha llegado, intenta nuevamente en " & AOG_RESEND_INTERVAL \ 10000 & " segundos")
                        
            End If
        
        Else
            
            ' No hay registros en `account_guard` = no tiene codigo = le mando uno
            EnviarCode = True
            
        End If
        
        If EnviarCode Then
        
            If TRANSPORT_METHOD = "API" Then
                Call GenerarCodigoAPI(UserIndex)
            Else
                Call GenerarCodigo(UserIndex)
            End If
    
        End If
        
    End With

    Exit Sub

EnviarCodigo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.EnviarCodigo", Erl)
    
End Sub

'---------------------------------------------------------------------------------------------------
' Si VerificarOrigen = False, le notificamos al usuario que ponga el codigo que le mandamos al mail.
'---------------------------------------------------------------------------------------------------
Public Sub WriteGuardNotice(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.GuardNotice)
        Call .EndPacket
        
        Call EnviarCodigo(UserIndex)
        
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
    
100     With UserList(UserIndex)
        
102         Dim Codigo As String: Codigo = .incomingData.ReadASCIIString

            If .AccountID = 0 Then Exit Sub

104         Call MakeQuery("SELECT TIMESTAMPDIFF(SECOND, `created_at`, CURRENT_TIMESTAMP) AS delta_time, code FROM account_guard WHERE account_id = ?", False, .AccountID)
        
            ' El codigo expira despues de 1 minuto.
106         If AOG_EXPIRE <> 0 And QueryData!delta_time > AOG_EXPIRE Then
            
                ' Le avisamos que expiro
108             Call WriteShowMessageBox(UserIndex, "El código de verificación ha expirado.")
110             Debug.Print "El codigo expiro. Se generara uno nuevo!"
            
                ' Invalidamos el codigo
112             Call MakeQuery("DELETE FROM account_guard WHERE account_id = ?", True, UserList(UserIndex).AccountID)
            
                ' Lo kickeamos.
114             Call CloseSocket(UserIndex)
                 
            Else ' El codigo NO expiro...
            
                ' Lo comparamos con lo que tenemos en la BD
116             If Codigo = QueryData!code Then
            
118                 Call WritePersonajesDeCuenta(UserIndex)
120                 Call WriteMostrarCuenta(UserIndex)
                
                    ' Invalidamos el codigo
122                 Call MakeQuery("DELETE FROM account_guard WHERE account_id = ?", True, UserList(UserIndex).AccountID)
                
                Else
            
                    ' Le avisamos
124                 Call WriteShowMessageBox(UserIndex, "El código de verificación ha incorrecto.")
                
                    ' Lo kickeamos.
126                 Call CloseSocket(UserIndex)
                
                End If
            
            End If
 
        End With
    
        Exit Sub

HandleGuardNoticeResponse_Err:
128     Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuardNoticeResponse", Erl)
130     Call UserList(UserIndex).incomingData.SafeClearPacket

End Sub

Public Sub HandleGuardResendVerificationCode(ByVal UserIndex As Integer)
        
    On Error GoTo HandleResendVerificationCode_Err:
        
    Call EnviarCodigo(UserIndex)
        
    Exit Sub

HandleResendVerificationCode_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuardResendVerificationCode", Erl)
    Call UserList(UserIndex).incomingData.SafeClearPacket
    
End Sub

Private Sub GenerarCodigoAPI(ByVal UserIndex As Integer)

        On Error GoTo GenerarCodigoAPI_Err
    
        '------------------------------------------------------
        ' Preparamos las cosas para hacer la peticion
        '------------------------------------------------------
        Dim client  As New MSXML2.ServerXMLHTTP60
    
        Dim request As New clsRequestHandler
100     Call request.Initialize(client)
    
        ' Seteamos un objeto para manejar la peticion async
102     client.OnReadyStateChange = request
    
        Dim Codigo As String
104     Codigo = RandomString(5)
        
106     Debug.Print Codigo
        '------------------------------------------------------
        ' Hacemos la peticion
        '------------------------------------------------------
108     client.Open "POST", API_ENDPOINT, True
    
110     client.setRequestHeader "x-api-key", API_KEY
112     client.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            
114     client.send "account_id=" & UserList(UserIndex).AccountID & _
           "&email=" & UserList(UserIndex).Cuenta & _
           "&ip_address=" & UserList(UserIndex).IP & _
           "&code=" & Codigo
    
        Exit Sub

GenerarCodigoAPI_Err:
        Call RegistrarError(Err.Number, Err.Description, "Protocol.GenerarCodigoAPI", Erl)
        
End Sub

Private Sub GenerarCodigo(ByVal UserIndex As Integer)

        On Error GoTo GenerarCodigo_Err
    
        Dim Codigo      As String
        Dim NuevoCodigo As Boolean
    
100     With UserList(UserIndex)
        
102         Call MakeQuery("SELECT TIMESTAMPDIFF(SECOND, `created_at`, CURRENT_TIMESTAMP) AS time_diff, code FROM account_guard WHERE account_id = ?", False, .AccountID)
     
            ' NO tiene codigo
104         If QueryData Is Nothing Then

106             NuevoCodigo = True
        
                ' Tiene codigo, pero ya expiro...
108         ElseIf AOG_EXPIRE <> 0 And QueryData!time_diff > AOG_EXPIRE Then
            
110             NuevoCodigo = True
            
            Else ' Si ya tiene codigo y NO expiro...
            
112             NuevoCodigo = False

            End If
        
114         If NuevoCodigo Then
        
                ' Generamos un nuevo codigo
116             Codigo = RandomString(5)
                  
                ' Lo guardamos en la BD
118             Call MakeQuery("REPLACE INTO account_guard (account_id, code) VALUES (?, ?)", True, .AccountID, Codigo)
        
            Else
            
                ' Usamos el codigo vigente
120             Codigo = QueryData!code
            
            End If
        
122         Debug.Print "Codigo de Verificacion: " & Codigo & vbNewLine
        
            ' Enviamos el mail con el codigo
124         Call SendEmail(UserList(UserIndex).Cuenta, Codigo, UserList(UserIndex).IP)

        End With
    
        '<EhFooter>
        Exit Sub

GenerarCodigo_Err:
        Call RegistrarError(Err.Number, Err.Description, "Protocol.GenerarCodigo", Erl)

End Sub

' Source: https://accautomation.ca/how-to-send-email-to-smtp-server/
Sub SendEmail(ByVal Email As String, ByVal Codigo As String, ByVal IP As String)

    On Error Resume Next
    
    If LenB(SMTP_HOST) = 0 Or _
       LenB(SMTP_PORT) = 0 Or _
       LenB(SMTP_USER) = 0 Or _
       LenB(SMTP_PASS) = 0 Then Exit Sub
    
    Dim Schema    As String
    
    Dim cdoMsg    As Object
    Dim cdoConf   As Object
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
           "El equipo de Noland Studios"
                    
        Set .Configuration = cdoConf
        
        ' Send the message
        Call .send

    End With

    'Check for errors and display message
    If Err.Number <> 0 Then
        Call RegistrarError(Err.Number, "Error al enviar correo a " & Email & vbNewLine & Err.Description, "AOGuard.SendMail")

    End If

    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set cdoFields = Nothing

End Sub
