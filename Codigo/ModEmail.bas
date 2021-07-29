Attribute VB_Name = "ModCuentas"
Option Explicit

Public Function EnviarCorreo(ByVal UserMail As String) As Boolean
        
        On Error GoTo EnviarCorreo_Err
        
        'Shell App.Path & "\cuentas.exe *" & UserMail & "*" & ObtenerCodigo(UserMail) & "*" ' & UserName
    
100     EnviarCorreo = True

        
        Exit Function

EnviarCorreo_Err:
102     Call TraceError(Err.Number, Err.Description, "ModCuentas.EnviarCorreo", Erl)

        
End Function

Public Function EnviarCorreoRecuperacion(ByVal UserNick As String, ByVal UserMail As String) As Boolean
        
        On Error GoTo EnviarCorreoRecuperacion_Err
        

100     If UserNick = "" Then
102         EnviarCorreoRecuperacion = False
            Exit Function

        End If

104     If UserMail = "" Then
106         EnviarCorreoRecuperacion = False
            Exit Function

        End If
    
        ' WyroX: Desactivo esto, porque ahora las contrasenias se hashean
        ' Hay que reveer el sistema
        'Shell App.Path & "\RecuperarPass.exe" & " " & UserNick & "*" & UserMail & "*" & SDesencriptar(ObtenerPASSWORD(UserNick))
108     EnviarCorreoRecuperacion = True

        
        Exit Function

EnviarCorreoRecuperacion_Err:
110     Call TraceError(Err.Number, Err.Description, "ModCuentas.EnviarCorreoRecuperacion", Erl)

        
End Function

Public Function ObtenerCodigo(ByVal Name As String) As String
        
        On Error GoTo ObtenerCodigo_Err
        

100     If Database_Enabled Then
102         ObtenerCodigo = GetCodigoActivacionDatabase(Name)
        Else
104         ObtenerCodigo = GetVar(CuentasPath & UCase$(Name) & ".act", "INIT", "CodigoActivacion")

        End If

        
        Exit Function

ObtenerCodigo_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerCodigo", Erl)

        
End Function

Public Function ObtenerValidacion(ByVal Name As String) As Boolean
        
        On Error GoTo ObtenerValidacion_Err
        

100     If Database_Enabled Then
102         ObtenerValidacion = CheckCuentaActivadaDatabase(Name)
        Else
104         ObtenerValidacion = val(GetVar(CuentasPath & UCase$(Name) & ".act", "INIT", "Activada"))

        End If
    
        
        Exit Function

ObtenerValidacion_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerValidacion", Erl)

        
End Function

Public Function ObtenerEmail(ByVal Name As String) As String
        
        On Error GoTo ObtenerEmail_Err
        

100     If Database_Enabled Then
102         ObtenerEmail = GetEmailDatabase(Name)
        Else
104         ObtenerEmail = GetVar(CuentasPath & UCase$(Name) & ".act", "INIT", "Email")

        End If
    
        
        Exit Function

ObtenerEmail_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerEmail", Erl)

        
End Function

Public Function ObtenerMacAdress(ByVal Name As String) As String
        
        On Error GoTo ObtenerMacAdress_Err
        

100     If Database_Enabled Then
102         ObtenerMacAdress = GetMacAddressDatabase(Name)
        Else
104         ObtenerMacAdress = GetVar(CuentasPath & UCase$(Name) & ".act", "INIT", "MacAdress")

        End If

        
        Exit Function

ObtenerMacAdress_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerMacAdress", Erl)

        
End Function

Public Function ObtenerHDserial(ByVal Name As String) As Long
        
        On Error GoTo ObtenerHDserial_Err
        

100     If Database_Enabled Then
102         ObtenerHDserial = GetHDSerialDatabase(Name)
        Else
104         ObtenerHDserial = GetVar(CuentasPath & UCase$(Name) & ".act", "INIT", "HDserial")

        End If

        
        Exit Function

ObtenerHDserial_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerHDserial", Erl)

        
End Function

Public Function CuentaExiste(ByVal CuentaEmail As String) As Boolean
        
        On Error GoTo CuentaExiste_Err
        

100     If Database_Enabled Then
102         CuentaExiste = CheckCuentaExiste(CuentaEmail)
        Else
104         CuentaExiste = FileExist(CuentasPath & LCase$(CuentaEmail) & ".act", vbNormal)

        End If

        
        Exit Function

CuentaExiste_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.CuentaExiste", Erl)

        
End Function

Public Function ObtenerCuenta(ByVal Name As String) As String
        
        On Error GoTo ObtenerCuenta_Err
        
102         ObtenerCuenta = GetNombreCuentaDatabase(Name)

    
        
        Exit Function

ObtenerCuenta_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerCuenta", Erl)

        
End Function

Public Function PasswordValida(Password As String, PasswordHash As String, Salt As String) As Boolean
        
        On Error GoTo PasswordValida_Err
        

        Dim oSHA256 As CSHA256

100     Set oSHA256 = New CSHA256

102     PasswordValida = (PasswordHash = oSHA256.SHA256(Password & Salt))
    
104     Set oSHA256 = Nothing

        
        Exit Function

PasswordValida_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.PasswordValida", Erl)

        
End Function

Public Function ObtenerBaneo(ByVal Name As String) As Boolean
        
        On Error GoTo ObtenerBaneo_Err
        

100     If Database_Enabled Then
102         ObtenerBaneo = CheckBanCuentaDatabase(Name)
        Else
104         ObtenerBaneo = val(GetVar(CuentasPath & LCase$(Name) & ".act", "BAN", "Baneada")) = 1

        End If

        
        Exit Function

ObtenerBaneo_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerBaneo", Erl)

        
End Function

Public Function ObtenerMotivoBaneo(ByVal Name As String) As String
        
        On Error GoTo ObtenerMotivoBaneo_Err
        

100     If Database_Enabled Then
102         ObtenerMotivoBaneo = GetMotivoBanCuentaDatabase(Name)
        Else
104         ObtenerMotivoBaneo = GetVar(CuentasPath & UCase$(Name) & ".act", "BAN", "Motivo")

        End If

        
        Exit Function

ObtenerMotivoBaneo_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerMotivoBaneo", Erl)

        
End Function

Public Function ObtenerQuienBaneo(ByVal Name As String) As String
        
        On Error GoTo ObtenerQuienBaneo_Err
        

100     If Database_Enabled Then
102         ObtenerQuienBaneo = GetQuienBanCuentaDatabase(Name)
        Else
104         ObtenerQuienBaneo = GetVar(CuentasPath & UCase$(Name) & ".act", "BAN", "BANEO")

        End If

        
        Exit Function

ObtenerQuienBaneo_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerQuienBaneo", Erl)

        
End Function

Public Function ObtenerCantidadDePersonajes(ByVal Name As String) As String
        
        On Error GoTo ObtenerCantidadDePersonajes_Err
        

100     If Database_Enabled Then
102         ObtenerCantidadDePersonajes = GetPersonajesCountDatabase(Name)
        Else
104         ObtenerCantidadDePersonajes = GetVar(CuentasPath & UCase$(Name) & ".act", "PERSONAJES", "Total")

        End If

        
        Exit Function

ObtenerCantidadDePersonajes_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerCantidadDePersonajes", Erl)

        
End Function

Public Function ObtenerCantidadDePersonajesByUserIndex(ByVal UserIndex As Integer) As Byte
        
        On Error GoTo ObtenerCantidadDePersonajesByUserIndex_Err
        

100     If Database_Enabled Then
102         ObtenerCantidadDePersonajesByUserIndex = GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID)
        Else
104         ObtenerCantidadDePersonajesByUserIndex = val(GetVar(CuentasPath & UCase$(UserList(UserIndex).Name) & ".act", "PERSONAJES", "Total"))

        End If

        
        Exit Function

ObtenerCantidadDePersonajesByUserIndex_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerCantidadDePersonajesByUserIndex", Erl)

        
End Function

Public Function ObtenerLogeada(ByVal Name As String) As Byte
        
        On Error GoTo ObtenerLogeada_Err
        

100     If Database_Enabled Then
102         ObtenerLogeada = GetCuentaLogeadaDatabase(Name)
        Else
104         ObtenerLogeada = GetVar(CuentasPath & UCase$(Name) & ".act", "INIT", "Logeada")

        End If

        
        Exit Function

ObtenerLogeada_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerLogeada", Erl)

        
End Function

Public Function ObtenerNombrePJ(ByVal Cuenta As String, ByVal i As Byte) As String
        
        On Error GoTo ObtenerNombrePJ_Err
        
100     ObtenerNombrePJ = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)

        
        Exit Function

ObtenerNombrePJ_Err:
102     Call TraceError(Err.Number, Err.Description, "ModCuentas.ObtenerNombrePJ", Erl)

        
End Function

Public Function GetUserGuildIndex(ByVal UserName As String) As Integer
        
        On Error GoTo GetUserGuildIndex_Err
        

        '***************************************************
        'Author: Juan Andres Dalmasso
        'Last Modification: 18/09/2018
        '18/09/2018 CHOTS: Checks database too
        '***************************************************
100     If InStrB(UserName, "\") <> 0 Then
102         UserName = Replace(UserName, "\", vbNullString)

        End If

104     If InStrB(UserName, "/") <> 0 Then
106         UserName = Replace(UserName, "/", vbNullString)

        End If

108     If InStrB(UserName, ".") <> 0 Then
110         UserName = Replace(UserName, ".", vbNullString)

        End If

116     GetUserGuildIndex = GetUserGuildIndexDatabase(UserName)

        Exit Function

GetUserGuildIndex_Err:
118     Call TraceError(Err.Number, Err.Description, "ModCuentas.GetUserGuildIndex", Erl)

End Function

Public Function ObtenerCriminal(ByVal Name As String) As Byte

        On Error GoTo ErrorHandler
    
        Dim Criminal As Byte

102     Criminal = GetUserStatusDatabase(Name)

106     If EsRolesMaster(Name) Then
108         Criminal = 3
110     ElseIf EsConsejero(Name) Then
112         Criminal = 4
114     ElseIf EsSemiDios(Name) Then
116         Criminal = 5
118     ElseIf EsDios(Name) Then
120         Criminal = 6
122     ElseIf EsAdmin(Name) Then
124         Criminal = 7
        End If

126     ObtenerCriminal = Criminal

        Exit Function
ErrorHandler:
128     ObtenerCriminal = 1

End Function

Public Function ObtenerMapa(ByVal Name As String) As String

        On Error GoTo ErrorHandler

        Dim Mapa As String

100     ObtenerMapa = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Position")
    
        Exit Function
ErrorHandler:
102     ObtenerMapa = "1-50-50"
    
End Function

Public Function ObtenerClase(ByVal Name As String) As Byte

        On Error GoTo ErrorHandler

100     ObtenerClase = GetVar(CharPath & UCase$(Name & ".chr"), "INIT", "Clase")

        Exit Function
ErrorHandler:
102     ObtenerClase = "1"

End Function
