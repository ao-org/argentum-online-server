Attribute VB_Name = "AOGuard"
Option Explicit

Private Const MAX_CODE_RESEND_COUNT As Byte = 10

' Configuracion - Argentum Guard
Public AOG_STATUS                   As Byte
Private AOG_EXPIRE                  As Long
Private AOG_RESEND_INTERVAL         As Integer
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
        
        On Error GoTo LoadAOGuardConfiguration_Err
        
100     If Not FileExist(IniPath & "AOGuard.ini", vbNormal) Then
102         AOG_STATUS = 0
            Exit Sub

        End If
    
        Dim ConfigFile As New clsIniManager
104     Call ConfigFile.Initialize(IniPath & "AOGuard.ini")
        
106     AOG_STATUS = val(ConfigFile.GetValue("INIT", "Enabled"))
108     AOG_EXPIRE = val(ConfigFile.GetValue("INIT", "CodeExpiresInSeconds"))
    
110     AOG_RESEND_INTERVAL = val(ConfigFile.GetValue("INIT", "CodeResendInterval"))

112     If AOG_RESEND_INTERVAL = 0 Then AOG_RESEND_INTERVAL = 50000
    
114     TRANSPORT_METHOD = UCase$(ConfigFile.GetValue("INIT", "TransportMethod"))

116     Select Case TRANSPORT_METHOD
    
            Case "API"
                ' Configuracion API
118             API_ENDPOINT = ConfigFile.GetValue("API", "Endpoint")
120             API_KEY = ConfigFile.GetValue("API", "Key")
            
122         Case "SMTP"
                ' Configuracion SMTP interno
                ' (no me gusta xq bloquea el hilo pero bueh, puede servir para salir del paso)
124             SMTP_HOST = ConfigFile.GetValue("SMTP", "HOST")
126             SMTP_PORT = val(ConfigFile.GetValue("SMTP", "PORT"))
128             SMTP_AUTH = val(ConfigFile.GetValue("SMTP", "AUTH"))
130             SMTP_SECURE = val(ConfigFile.GetValue("SMTP", "SECURE"))
132             SMTP_USER = ConfigFile.GetValue("SMTP", "USER")
134             SMTP_PASS = ConfigFile.GetValue("SMTP", "PASS")
            
        End Select
    
136     Set ConfigFile = Nothing

        Exit Sub

LoadAOGuardConfiguration_Err:
     Call TraceError(Err.Number, Err.Description, "AOGuard.LoadAOGuardConfiguration", Erl)

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
112     Call TraceError(Err.Number, Err.Description, "AOGuard.VerificarOrigen", Erl)

End Function

'---------------------------------------------------------------------------------------------------
' Le enviamos el codigo de verificacion al usuario si la situacion lo requiere
'---------------------------------------------------------------------------------------------------

Public Sub HandleNoticeResponse(ByVal UserIndex As Integer, ByVal Codigo As String)
        On Error GoTo HandleNoticeResponse_Err
    

100     With UserList(UserIndex)

104         If .AccountID = 0 Then Exit Sub

106         Call MakeQuery("SELECT TIMESTAMPDIFF(SECOND, `created_at`, CURRENT_TIMESTAMP) AS delta_time, code FROM account_guard WHERE account_id = ?", False, .AccountID)
        
            ' El codigo expira despues de 1 minuto.
108         If AOG_EXPIRE <> 0 And QueryData!delta_time > AOG_EXPIRE Then
            
                ' Le avisamos que expiro
110             Call WriteShowMessageBox(UserIndex, "El código de verificación ha expirado.")
112             Debug.Print "El codigo expiro. Se generara uno nuevo!"
            
                ' Invalidamos el codigo
114             Call MakeQuery("DELETE FROM account_guard WHERE account_id = ?", True, UserList(UserIndex).AccountID)
            
                ' Lo kickeamos.
116             Call CloseSocket(UserIndex)
                 
            Else ' El codigo NO expiro...
            
                ' Lo comparamos con lo que tenemos en la BD
118             If Codigo = QueryData!Code Then
            
120                 Call WritePersonajesDeCuenta(UserIndex)
122                 Call WriteMostrarCuenta(UserIndex)
                
                    ' Invalidamos el codigo
124                 Call MakeQuery("DELETE FROM account_guard WHERE account_id = ?", True, UserList(UserIndex).AccountID)
                
                Else
            
                    ' Le avisamos
126                 Call WriteShowMessageBox(UserIndex, "El código de verificación ha incorrecto.")
                
                    ' Lo kickeamos.
128                 Call CloseSocket(UserIndex)
                
                End If
            
            End If
 
        End With

HandleNoticeResponse_Err:
136     Call TraceError(Err.Number, Err.Description, "AOGuard.HandleNoticeResponse", Erl)
    
End Sub

Public Sub EnviarCodigo(ByVal UserIndex As Integer)

        On Error GoTo EnviarCodigo_Err
    
        Dim EnviarCode As Boolean
    
100     With UserList(UserIndex)
            
102         Call MakeQuery("SELECT TIMESTAMPDIFF(SECOND, `code_last_sent`, CURRENT_TIMESTAMP) AS delta_time, code_resend_attempts FROM account_guard WHERE account_id = ?", False, .AccountID)
        
            ' Hay registros en `account_guard` = tiene codigo = me fijo si le mando o no
104         If Not QueryData Is Nothing Then
            
                ' Ya te dije X veces que esperes un toque! Si no lo haces, sos alto bot!
106             If val(QueryData!code_resend_attempts) > MAX_CODE_RESEND_COUNT Then
108                 EnviarCode = False
110                 Call CloseSocket(UserIndex)
                    Exit Sub
                    
                End If
                
                ' Establecemos un intervalo de tiempo para volver a mandarle el codigo al usuario
112             If QueryData!delta_time > AOG_RESEND_INTERVAL Then
                    
114                 EnviarCode = True
                    
116                 Call MakeQuery("UPDATE account_guard SET code_last_sent = CURRENT_TIMESTAMP, code_resend_attempts = 0 WHERE account_id = ?", True, .AccountID)
                        
118                 Call WriteShowMessageBox(UserIndex, "Te hemos enviado un correo con el código de verificacion a tu correo. " & _
                                                        "Si no lo encuentras, revisa la carpeta de SPAM. " & _
                                                        "Si no te ha llegado, intenta nuevamente en " & val(QueryData!delta_time) & " segundos")
                        
                Else
                
120                 EnviarCode = False
                
122                 Call MakeQuery("UPDATE account_guard SET code_resend_attempts = code_resend_attempts + 1 WHERE account_id = ?", True, .AccountID)
                    
124                 Call WriteShowMessageBox(UserIndex, "Ya te hemos enviado un correo con el código de verificacion. " & _
                                                        "Si no te ha llegado, intenta nuevamente en " & val(QueryData!delta_time) & " segundos")
                        
                End If
        
            Else
            
                ' No hay registros en `account_guard` = no tiene codigo = le mando uno
126             EnviarCode = True
            
            End If
        
128         If EnviarCode Then
        
130             If TRANSPORT_METHOD = "API" Then
132                 Call GenerarCodigoAPI(UserIndex)
                Else
134                 Call GenerarCodigo(UserIndex)
                End If
    
            End If
        
        End With

        Exit Sub

EnviarCodigo_Err:
136     Call TraceError(Err.Number, Err.Description, "AOGuard.EnviarCodigo", Erl)
    
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
116     Call TraceError(Err.Number, Err.Description, "AOGuard.GenerarCodigoAPI", Erl)
        
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
        Call TraceError(Err.Number, Err.Description, "AOGuard.GenerarCodigo", Erl)

End Sub

' Source: https://accautomation.ca/how-to-send-email-to-smtp-server/
Sub SendEmail(ByVal Email As String, ByVal Codigo As String, ByVal IP As String)

        On Error Resume Next
    
100     If LenB(SMTP_HOST) = 0 Or _
           LenB(SMTP_PORT) = 0 Or _
           LenB(SMTP_USER) = 0 Or _
           LenB(SMTP_PASS) = 0 Then Exit Sub
    
        Dim Schema    As String
    
        Dim cdoMsg    As Object
        Dim cdoConf   As Object
        Dim cdoFields As Object
    
102     Set cdoMsg = CreateObject("CDO.Message")
104     Set cdoConf = CreateObject("CDO.Configuration")
106     Set cdoFields = cdoConf.Fields
    
        ' Send one copy with Google SMTP server (with autentication)
108     Schema = "http://schemas.microsoft.com/cdo/configuration/"
    
110     cdoFields.Item(Schema & "sendusing") = 2
112     cdoFields.Item(Schema & "smtpserver") = SMTP_HOST
114     cdoFields.Item(Schema & "smtpserverport") = SMTP_PORT
116     cdoFields.Item(Schema & "smtpauthenticate") = SMTP_AUTH
118     cdoFields.Item(Schema & "sendusername") = SMTP_USER
120     cdoFields.Item(Schema & "sendpassword") = SMTP_PASS
122     cdoFields.Item(Schema & "smtpusessl") = SMTP_SECURE
    
124     Call cdoFields.Update

126     With cdoMsg
    
128         .To = Email
130         .From = "guardian@ao20.com.ar"
132         .Subject = "Argentum Guard - Acceso desde un nuevo dispositivo"
        
            ' Body of message can be any HTML code
134         .HTMLBody = "Hemos detectado un intento de acceso a tu cuenta desde un dispositivo desconocido <br /><br />" & _
               "IP: " & IP & "<br /><br />" & _
               "Si fuiste tu, te aparecerá un dialogo donde tendrás que ingresar el siguiente código: <strong>" & Codigo & "</strong>" & "<br />" & _
               "<strong>Si NO fuiste tu, ignora este mensaje y considera cambiar tu contraseña</strong>" & "<br /><br />" & _
               "El equipo de Noland Studios"
                    
136         Set .Configuration = cdoConf
        
            ' Send the message
138         Call .send

        End With

        'Check for errors and display message
140     If Err.Number <> 0 Then
142         Call TraceError(Err.Number, "Error al enviar correo a " & Email & vbNewLine & Err.Description, "AOGuard.SendMail")

        End If

144     Set cdoMsg = Nothing
146     Set cdoConf = Nothing
148     Set cdoFields = Nothing

End Sub
