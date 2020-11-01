Attribute VB_Name = "ModCuentas"
Option Explicit

Public Function EnviarCorreo(ByVal UserMail As String) As Boolean
    Shell App.Path & "\cuentas.exe *" & UserMail & "*" & ObtenerCodigo(UserMail) & "*" ' & UserName
    
    EnviarCorreo = True
End Function

Public Function EnviarCorreoRecuperacion(ByVal UserNick As String, ByVal UserMail As String) As Boolean
    If UserNick = "" Then
        EnviarCorreoRecuperacion = False
        Exit Function
    End If
    If UserMail = "" Then
        EnviarCorreoRecuperacion = False
        Exit Function
    End If
    
    ' WyroX: Desactivo esto, porque ahora las contrasenias se hashean
    ' Hay que reveer el sistema
    'Shell App.Path & "\RecuperarPass.exe" & " " & UserNick & "*" & UserMail & "*" & SDesencriptar(ObtenerPASSWORD(UserNick))
    EnviarCorreoRecuperacion = True
End Function

Public Function ObtenerCodigo(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerCodigo = GetCodigoActivacionDatabase(name)
    Else
        ObtenerCodigo = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "CodigoActivacion")
    End If

End Function

Public Function ObtenerValidacion(ByVal name As String) As Boolean

    If Database_Enabled Then
        ObtenerValidacion = CheckCuentaActivadaDatabase(name)
    Else
        ObtenerValidacion = val(GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "Activada"))
    End If
    
End Function

Public Function ObtenerEmail(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerEmail = GetEmailDatabase(name)
    Else
        ObtenerEmail = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "Email")
    End If
    
End Function

Public Function ObtenerMacAdress(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerMacAdress = GetMacAddressDatabase(name)
    Else
        ObtenerMacAdress = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "MacAdress")
    End If

End Function

Public Function ObtenerHDserial(ByVal name As String) As Long

    If Database_Enabled Then
        ObtenerHDserial = GetHDSerialDatabase(name)
    Else
        ObtenerHDserial = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "HDserial")
    End If

End Function

Public Function CuentaExiste(ByVal CuentaEmail As String) As Boolean

    If Database_Enabled Then
        CuentaExiste = CheckCuentaExiste(CuentaEmail)
    Else
        CuentaExiste = FileExist(CuentasPath & LCase$(CuentaEmail) & ".act", vbNormal)
    End If

End Function

Public Sub SaveNewAccount(ByVal UserIndex As Integer, ByVal CuentaEmail As String, ByVal Password As String)

    Dim Salt As String * 10
    Salt = RandomString(10) ' Alfanumerico
    
    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256

    Dim PasswordHash As String * 64
    PasswordHash = oSHA256.SHA256(Password & Salt)
    
    Set oSHA256 = Nothing

    Dim Codigo As String * 6
    Codigo = RandomString(6, True) ' Letras mayusculas y numeros

    If Database_Enabled Then
        Call SaveNewAccountDatabase(CuentaEmail, PasswordHash, Salt, Codigo)
    Else
        Call SaveNewAccountCharfile(CuentaEmail, PasswordHash, Salt, Codigo)
    End If

End Sub

Public Sub SaveNewAccountCharfile(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)
On Error GoTo ErrorHandler

    Dim Manager     As clsIniReader
    
    Dim AccountFile As String

    Set Manager = New clsIniReader
    AccountFile = CuentasPath & LCase$(CuentaEmail) & ".act"

    With Manager
        
        Call .ChangeValue("INIT", "Email", CuentaEmail)
        Call .ChangeValue("INIT", "Password", PasswordHash)
        Call .ChangeValue("INIT", "Salt", Salt)
        Call .ChangeValue("INIT", "Activada", "0")
        Call .ChangeValue("INIT", "FechaCreacion", Date)
        Call .ChangeValue("INIT", "CodigoActivacion", Codigo)
        Call .ChangeValue("PERSONAJES", "Total", "0")
        Call .ChangeValue("INIT", "Logeada", "0")
        Call .ChangeValue("BAN", "Baneada", "0")
        Call .ChangeValue("BAN", "Motivo", "")
        Call .ChangeValue("BAN", "BANEO", "")
        
        'Grabamos donador
        Call .ChangeValue("DONADOR", "DONADOR", "0")
        Call .ChangeValue("DONADOR", "CREDITOS", "0")
        Call .ChangeValue("DONADOR", "FECHAEXPIRACION", "")
        
        'Seguridad Ladder
        Call .ChangeValue("INIT", "MacAdress", "0")
        Call .ChangeValue("INIT", "HDserial", "0")

        
        Call .DumpFile(AccountFile)

    End With

    Set Manager = Nothing

    Exit Sub

ErrorHandler:
    Call LogError("Error en SaveNewAccountCharfile. ")
End Sub

Public Function ObtenerCuenta(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerCuenta = GetNombreCuentaDatabase(name)
    Else
        ObtenerCuenta = GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Cuenta")
    End If
    
End Function

Public Function PasswordValida(Cuenta As String, Password As String) As Boolean

    Dim PasswordHash As String * 64
    Dim Salt As String * 10

    If Database_Enabled Then
        Call GetPasswordAndSaltDatabase(Cuenta, PasswordHash, Salt)
    Else
        PasswordHash = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "INIT", "PASSWORD")
        Salt = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "INIT", "SALT")
    End If
    
    Dim oSHA256 As CSHA256

    Set oSHA256 = New CSHA256

    PasswordValida = (PasswordHash = oSHA256.SHA256(Password & Salt))
    
    Set oSHA256 = Nothing

End Function

Public Function ObtenerBaneo(ByVal name As String) As Boolean

    If Database_Enabled Then
        ObtenerBaneo = CheckBanCuentaDatabase(name)
    Else
        ObtenerBaneo = val(GetVar(CuentasPath & LCase$(name) & ".act", "BAN", "Baneada")) = 1
    End If

End Function

Public Function ObtenerMotivoBaneo(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerMotivoBaneo = GetMotivoBanCuentaDatabase(name)
    Else
        ObtenerMotivoBaneo = GetVar(CuentasPath & UCase$(name) & ".act", "BAN", "Motivo")
    End If

End Function

Public Function ObtenerQuienBaneo(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerQuienBaneo = GetQuienBanCuentaDatabase(name)
    Else
        ObtenerQuienBaneo = GetVar(CuentasPath & UCase$(name) & ".act", "BAN", "BANEO")
    End If

End Function

Public Function ObtenerCantidadDePersonajes(ByVal name As String) As String

    If Database_Enabled Then
        ObtenerCantidadDePersonajes = GetPersonajesCountDatabase(name)
    Else
        ObtenerCantidadDePersonajes = GetVar(CuentasPath & UCase$(name) & ".act", "PERSONAJES", "Total")
    End If

End Function

Public Function ObtenerCantidadDePersonajesByUserIndex(ByVal UserIndex As Integer) As Byte

    If Database_Enabled Then
        ObtenerCantidadDePersonajesByUserIndex = GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID)
    Else
        ObtenerCantidadDePersonajesByUserIndex = val(GetVar(CuentasPath & UCase$(UserList(UserIndex).name) & ".act", "PERSONAJES", "Total"))
    End If

End Function

Public Function ObtenerLogeada(ByVal name As String) As Byte

    If Database_Enabled Then
        ObtenerLogeada = GetCuentaLogeadaDatabase(name)
    Else
        ObtenerLogeada = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "Logeada")
    End If

End Function

Public Function ObtenerNombrePJ(ByVal Cuenta As String, ByVal i As Byte) As String
    ObtenerNombrePJ = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)
End Function

Sub BorrarPJdeCuenta(ByVal name As String)
    Dim CantpjsNew As Byte
    Dim CantpjsOld As Byte
    Dim indice As Byte
    Dim pjs(1 To 8) As String
    Dim SiguientePJ As Byte
    Dim Cuenta As String

    Cuenta = ObtenerCuenta(name)
    
    CantpjsOld = ObtenerCantidadDePersonajes(Cuenta)
    
    Dim i As Integer
    For i = 1 To CantpjsOld
        pjs(i) = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)
        If pjs(i) = name Then
            indice = i
            pjs(i) = ""
        End If
    Next i
    
    Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & indice, "")
    
    
    For i = 1 To CantpjsOld
        pjs(i) = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)
    Next i
    
    For i = 1 To CantpjsOld
        If pjs(i) = "" And i + 1 < 9 Then
            pjs(i) = pjs(i + 1)
            pjs(i + 1) = ""
        Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i, pjs(i))
        Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i + 1, "")
        End If
    Next i
    
    SiguientePJ = ObtenerCantidadDePersonajes(Cuenta) - 1
    Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "Total", SiguientePJ)

End Sub

Sub GrabarNuevoPjEnCuentaCharfile(ByVal UserCuenta As String, ByVal name As String)

    Dim cantidaddePersonajes As Byte
    cantidaddePersonajes = ObtenerCantidadDePersonajes(UserCuenta)

    Call WriteVar(CuentasPath & UCase$(UserCuenta) & ".act", "PERSONAJES", "Total", cantidaddePersonajes + 1)
    Call WriteVar(CuentasPath & UCase$(UserCuenta) & ".act", "PERSONAJES", "PJ" & cantidaddePersonajes + 1, name)

End Sub

Sub BorrarCuenta(ByVal CuentaName As String)

    If Database_Enabled Then
        Call BorrarCuentaDatabase(CuentaName)
    Else
        'Primero borramos PJ POR PJ, se copia los personajes a la carpeta de personajes borrados
        Dim Cantpjs As Byte
        
        Cantpjs = ObtenerCantidadDePersonajes(CuentaName)
        Dim indice As Byte
        Dim pjs(1 To 8) As String
        
        Dim i As Integer
        For i = 1 To Cantpjs
            pjs(i) = GetVar(CuentasPath & UCase$(CuentaName) & ".act", "PERSONAJES", "PJ" & i)
            If FileExist(CharPath & UCase$(pjs(i)) & ".chr", vbNormal) Then
                Call FileCopy(CharPath & UCase$(pjs(i)) & ".chr", DeletePath & UCase$(pjs(i)) & ".chr")
                Call Kill(CharPath & UCase$(pjs(i)) & ".chr")
            End If
        Next i
    
        If FileExist(CuentasPath & UCase$(CuentaName) & ".act", vbNormal) Then
            Call FileCopy(CuentasPath & UCase$(CuentaName) & ".act", DeleteCuentasPath & UCase$(CuentaName))
            Call Kill(CuentasPath & CuentaName & ".act")
        End If
    End If

End Sub

Public Function ObtenerNivel(ByVal name As String) As Byte
On Error GoTo ErrorHandler
ObtenerNivel = GetVar(CharPath & UCase$(name & ".chr"), "STATS", "ELV")

Exit Function
ErrorHandler:
ObtenerNivel = 1
End Function
Public Function ObtenerCuerpo(ByVal name As String) As Integer

On Error GoTo ErrorHandler
Dim EstaMuerto As Byte
Dim cuerpo As Long

EstaMuerto = GetVar(CharPath & UCase$(name & ".chr"), "flags", "Muerto")
If EstaMuerto = 0 Then
cuerpo = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Body")
    ObtenerCuerpo = cuerpo
Else
    ObtenerCuerpo = iCuerpoMuerto
End If

Exit Function
ErrorHandler:
ObtenerCuerpo = 1

End Function
Public Function ObtenerCabeza(ByVal name As String) As Integer
On Error GoTo ErrorHandler
Dim Head As String
Dim EstaMuerto As Byte

EstaMuerto = GetVar(CharPath & UCase$(name & ".chr"), "flags", "Muerto")

If EstaMuerto = 0 Then
    Head = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Head")
Else
    Head = iCabezaMuerto
End If

ObtenerCabeza = Head


Exit Function
ErrorHandler:
ObtenerCabeza = 1

End Function
Public Function ObtenerEscudo(ByVal name As String) As Integer
On Error GoTo ErrorHandler
ObtenerEscudo = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Escudo")
Exit Function
ErrorHandler:
ObtenerEscudo = 0
End Function
Public Function ObtenerArma(ByVal name As String) As Integer
On Error GoTo ErrorHandler
ObtenerArma = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Arma")
Exit Function
ErrorHandler:
ObtenerArma = 0
End Function
Public Function ObtenerCasco(ByVal name As String) As Integer
On Error GoTo ErrorHandler
ObtenerCasco = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Casco")
Exit Function
ErrorHandler:
ObtenerCasco = 0
End Function

Public Function GetUserGuildIndex(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    If InStrB(UserName, "\") <> 0 Then
        UserName = Replace(UserName, "\", vbNullString)

    End If

    If InStrB(UserName, "/") <> 0 Then
        UserName = Replace(UserName, "/", vbNullString)

    End If

    If InStrB(UserName, ".") <> 0 Then
        UserName = Replace(UserName, ".", vbNullString)

    End If

    If Not Database_Enabled Then
        GetUserGuildIndex = GetUserGuildIndexCharfile(UserName)
    Else
        GetUserGuildIndex = GetUserGuildIndexDatabase(UserName)

    End If

End Function

Public Function GetUserGuildIndexCharfile(ByRef UserName As String) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/09/2018
    '26/09/2018 CHOTS: Moved to FileIO
    '***************************************************
    Dim Temps As String
    
    Temps = GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX")

    If IsNumeric(Temps) Then
        GetUserGuildIndexCharfile = CInt(Temps)
    Else
        GetUserGuildIndexCharfile = 0

    End If

End Function

Public Function GetUserGuildPedidosCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildPedidosCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "Pedidos")

End Function

Sub SaveUserGuildPedidosCharfile(ByVal UserName As String, ByVal Pedidos As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "Pedidos", Pedidos)

End Sub

Sub SaveUserGuildMemberCharfile(ByVal UserName As String, ByVal guilds As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "Miembro", guilds)

End Sub

Sub SaveUserGuildIndexCharfile(ByVal UserName As String, ByVal GuildIndex As Integer)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX", GuildIndex)

End Sub

Sub SaveUserGuildAspirantCharfile(ByVal UserName As String, _
                                  ByVal AspirantIndex As Integer)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "ASPIRANTEA", AspirantIndex)

End Sub

Sub SendCharacterInfoCharfile(ByVal UserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************
    Dim gName       As String

    Dim UserFile    As clsIniReader

    Dim Miembro     As String

    Dim GuildActual As Integer

    ' Get the character's current guild
    GuildActual = GetUserGuildIndex(UserName)

    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"

    End If
    
    'Get previous guilds
    Miembro = GetUserGuildMember(UserName)

    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)

    End If

    Set UserFile = New clsIniReader

    With UserFile
        .Initialize (CharPath & UserName & ".chr")
    
        Call WriteCharacterInfo(UserIndex, UserName, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), .GetValue("STATS", "Banco"), .GetValue("GUILD", "Pedidos"), gName, Miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))

    End With
    
    Set UserFile = Nothing

End Sub

Public Function GetUserGuildMemberCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildMemberCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "Miembro")

End Function

Public Function GetUserGuildAspirantCharfile(ByVal UserName As String) As Integer
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildAspirantCharfile = val(GetVar(CharPath & UserName & ".chr", "GUILD", "ASPIRANTEA"))

End Function

Public Function GetUserGuildRejectionReasonCharfile(ByVal UserName As String) As String
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    GetUserGuildRejectionReasonCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "MotivoRechazo")

End Function

Sub SaveUserGuildRejectionReasonCharfile(ByVal UserName As String, ByVal Reason As String)
    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    '***************************************************

    Call WriteVar(CharPath & UserName & ".chr", "GUILD", "MotivoRechazo", Reason)

End Sub

Public Function ObtenerCriminal(ByVal name As String) As Byte

On Error GoTo ErrorHandler
    
    Dim Criminal As Byte

    If Database_Enabled Then
        Criminal = GetUserStatusDatabase(name)
    Else
        Criminal = GetVar(CharPath & UCase$(name & ".chr"), "FACCIONES", "Status")
    End If

    If EsRolesMaster(name) Then
        Criminal = 3
    ElseIf EsConsejero(name) Then
        Criminal = 4
    ElseIf EsSemiDios(name) Then
        Criminal = 5
    ElseIf EsDios(name) Then
        Criminal = 6
    ElseIf EsAdmin(name) Then
        Criminal = 7
    End If


    ObtenerCriminal = Criminal


Exit Function
ErrorHandler:
ObtenerCriminal = 1


End Function
Public Function ObtenerMapa(ByVal name As String) As String

On Error GoTo ErrorHandler

Dim Mapa As String

    ObtenerMapa = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Position")
    
    
    Exit Function
ErrorHandler:
ObtenerMapa = "1-50-50"

    
End Function
Public Function ObtenerClase(ByVal name As String) As Byte

On Error GoTo ErrorHandler

ObtenerClase = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Clase")

Exit Function
ErrorHandler:
ObtenerClase = "1"

End Function
