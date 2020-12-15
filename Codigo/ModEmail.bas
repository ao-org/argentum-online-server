Attribute VB_Name = "ModCuentas"
Option Explicit

Public Function EnviarCorreo(ByVal UserMail As String) As Boolean
        
        On Error GoTo EnviarCorreo_Err
        
100     Shell App.Path & "\cuentas.exe *" & UserMail & "*" & ObtenerCodigo(UserMail) & "*" ' & UserName
    
102     EnviarCorreo = True

        
        Exit Function

EnviarCorreo_Err:
104     Call RegistrarError(Err.Number, Err.description, "ModCuentas.EnviarCorreo", Erl)
106     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "ModCuentas.EnviarCorreoRecuperacion", Erl)
112     Resume Next
        
End Function

Public Function ObtenerCodigo(ByVal name As String) As String
        
        On Error GoTo ObtenerCodigo_Err
        

100     If Database_Enabled Then
102         ObtenerCodigo = GetCodigoActivacionDatabase(name)
        Else
104         ObtenerCodigo = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "CodigoActivacion")

        End If

        
        Exit Function

ObtenerCodigo_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerCodigo", Erl)
108     Resume Next
        
End Function

Public Function ObtenerValidacion(ByVal name As String) As Boolean
        
        On Error GoTo ObtenerValidacion_Err
        

100     If Database_Enabled Then
102         ObtenerValidacion = CheckCuentaActivadaDatabase(name)
        Else
104         ObtenerValidacion = val(GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "Activada"))

        End If
    
        
        Exit Function

ObtenerValidacion_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerValidacion", Erl)
108     Resume Next
        
End Function

Public Function ObtenerEmail(ByVal name As String) As String
        
        On Error GoTo ObtenerEmail_Err
        

100     If Database_Enabled Then
102         ObtenerEmail = GetEmailDatabase(name)
        Else
104         ObtenerEmail = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "Email")

        End If
    
        
        Exit Function

ObtenerEmail_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerEmail", Erl)
108     Resume Next
        
End Function

Public Function ObtenerMacAdress(ByVal name As String) As String
        
        On Error GoTo ObtenerMacAdress_Err
        

100     If Database_Enabled Then
102         ObtenerMacAdress = GetMacAddressDatabase(name)
        Else
104         ObtenerMacAdress = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "MacAdress")

        End If

        
        Exit Function

ObtenerMacAdress_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerMacAdress", Erl)
108     Resume Next
        
End Function

Public Function ObtenerHDserial(ByVal name As String) As Long
        
        On Error GoTo ObtenerHDserial_Err
        

100     If Database_Enabled Then
102         ObtenerHDserial = GetHDSerialDatabase(name)
        Else
104         ObtenerHDserial = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "HDserial")

        End If

        
        Exit Function

ObtenerHDserial_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerHDserial", Erl)
108     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.CuentaExiste", Erl)
108     Resume Next
        
End Function

Public Sub SaveNewAccount(ByVal UserIndex As Integer, ByVal CuentaEmail As String, ByVal Password As String)
        
        On Error GoTo SaveNewAccount_Err
        

        Dim Salt As String * 10

100     Salt = RandomString(10) ' Alfanumerico
    
        Dim oSHA256 As CSHA256

102     Set oSHA256 = New CSHA256

        Dim PasswordHash As String * 64

104     PasswordHash = oSHA256.SHA256(Password & Salt)
    
106     Set oSHA256 = Nothing

        Dim Codigo As String * 6

108     Codigo = RandomString(6, True) ' Letras mayusculas y numeros

110     If Database_Enabled Then
112         Call SaveNewAccountDatabase(CuentaEmail, PasswordHash, Salt, Codigo)
        Else
114         Call SaveNewAccountCharfile(CuentaEmail, PasswordHash, Salt, Codigo)

        End If

        
        Exit Sub

SaveNewAccount_Err:
116     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveNewAccount", Erl)
118     Resume Next
        
End Sub

Public Sub SaveNewAccountCharfile(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)

        On Error GoTo ErrorHandler

        Dim Manager     As clsIniReader
    
        Dim AccountFile As String

100     Set Manager = New clsIniReader
102     AccountFile = CuentasPath & LCase$(CuentaEmail) & ".act"

104     With Manager
        
106         Call .ChangeValue("INIT", "Email", CuentaEmail)
108         Call .ChangeValue("INIT", "Password", PasswordHash)
110         Call .ChangeValue("INIT", "Salt", Salt)
112         Call .ChangeValue("INIT", "Activada", "0")
114         Call .ChangeValue("INIT", "FechaCreacion", Date)
116         Call .ChangeValue("INIT", "CodigoActivacion", Codigo)
118         Call .ChangeValue("PERSONAJES", "Total", "0")
120         Call .ChangeValue("INIT", "Logeada", "0")
122         Call .ChangeValue("BAN", "Baneada", "0")
124         Call .ChangeValue("BAN", "Motivo", "")
126         Call .ChangeValue("BAN", "BANEO", "")
        
            'Grabamos donador
128         Call .ChangeValue("DONADOR", "DONADOR", "0")
130         Call .ChangeValue("DONADOR", "CREDITOS", "0")
132         Call .ChangeValue("DONADOR", "FECHAEXPIRACION", "")
        
            'Seguridad Ladder
134         Call .ChangeValue("INIT", "MacAdress", "0")
136         Call .ChangeValue("INIT", "HDserial", "0")
        
138         Call .DumpFile(AccountFile)

        End With

140     Set Manager = Nothing

        Exit Sub

ErrorHandler:
142     Call LogError("Error en SaveNewAccountCharfile. ")

End Sub

Public Function ObtenerCuenta(ByVal name As String) As String
        
        On Error GoTo ObtenerCuenta_Err
        

100     If Database_Enabled Then
102         ObtenerCuenta = GetNombreCuentaDatabase(name)
        Else
104         ObtenerCuenta = GetVar(CharPath & UCase$(name) & ".chr", "INIT", "Cuenta")

        End If
    
        
        Exit Function

ObtenerCuenta_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerCuenta", Erl)
108     Resume Next
        
End Function

Public Function PasswordValida(Password As String, PasswordHash As String, Salt As String) As Boolean
        
        On Error GoTo PasswordValida_Err
        

        Dim oSHA256 As CSHA256

100     Set oSHA256 = New CSHA256

102     PasswordValida = (PasswordHash = oSHA256.SHA256(Password & Salt))
    
104     Set oSHA256 = Nothing

        
        Exit Function

PasswordValida_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.PasswordValida", Erl)
108     Resume Next
        
End Function

Public Function ObtenerBaneo(ByVal name As String) As Boolean
        
        On Error GoTo ObtenerBaneo_Err
        

100     If Database_Enabled Then
102         ObtenerBaneo = CheckBanCuentaDatabase(name)
        Else
104         ObtenerBaneo = val(GetVar(CuentasPath & LCase$(name) & ".act", "BAN", "Baneada")) = 1

        End If

        
        Exit Function

ObtenerBaneo_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerBaneo", Erl)
108     Resume Next
        
End Function

Public Function ObtenerMotivoBaneo(ByVal name As String) As String
        
        On Error GoTo ObtenerMotivoBaneo_Err
        

100     If Database_Enabled Then
102         ObtenerMotivoBaneo = GetMotivoBanCuentaDatabase(name)
        Else
104         ObtenerMotivoBaneo = GetVar(CuentasPath & UCase$(name) & ".act", "BAN", "Motivo")

        End If

        
        Exit Function

ObtenerMotivoBaneo_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerMotivoBaneo", Erl)
108     Resume Next
        
End Function

Public Function ObtenerQuienBaneo(ByVal name As String) As String
        
        On Error GoTo ObtenerQuienBaneo_Err
        

100     If Database_Enabled Then
102         ObtenerQuienBaneo = GetQuienBanCuentaDatabase(name)
        Else
104         ObtenerQuienBaneo = GetVar(CuentasPath & UCase$(name) & ".act", "BAN", "BANEO")

        End If

        
        Exit Function

ObtenerQuienBaneo_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerQuienBaneo", Erl)
108     Resume Next
        
End Function

Public Function ObtenerCantidadDePersonajes(ByVal name As String) As String
        
        On Error GoTo ObtenerCantidadDePersonajes_Err
        

100     If Database_Enabled Then
102         ObtenerCantidadDePersonajes = GetPersonajesCountDatabase(name)
        Else
104         ObtenerCantidadDePersonajes = GetVar(CuentasPath & UCase$(name) & ".act", "PERSONAJES", "Total")

        End If

        
        Exit Function

ObtenerCantidadDePersonajes_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerCantidadDePersonajes", Erl)
108     Resume Next
        
End Function

Public Function ObtenerCantidadDePersonajesByUserIndex(ByVal UserIndex As Integer) As Byte
        
        On Error GoTo ObtenerCantidadDePersonajesByUserIndex_Err
        

100     If Database_Enabled Then
102         ObtenerCantidadDePersonajesByUserIndex = GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID)
        Else
104         ObtenerCantidadDePersonajesByUserIndex = val(GetVar(CuentasPath & UCase$(UserList(UserIndex).name) & ".act", "PERSONAJES", "Total"))

        End If

        
        Exit Function

ObtenerCantidadDePersonajesByUserIndex_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerCantidadDePersonajesByUserIndex", Erl)
108     Resume Next
        
End Function

Public Function ObtenerLogeada(ByVal name As String) As Byte
        
        On Error GoTo ObtenerLogeada_Err
        

100     If Database_Enabled Then
102         ObtenerLogeada = GetCuentaLogeadaDatabase(name)
        Else
104         ObtenerLogeada = GetVar(CuentasPath & UCase$(name) & ".act", "INIT", "Logeada")

        End If

        
        Exit Function

ObtenerLogeada_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerLogeada", Erl)
108     Resume Next
        
End Function

Public Function ObtenerNombrePJ(ByVal Cuenta As String, ByVal i As Byte) As String
        
        On Error GoTo ObtenerNombrePJ_Err
        
100     ObtenerNombrePJ = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)

        
        Exit Function

ObtenerNombrePJ_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.ObtenerNombrePJ", Erl)
104     Resume Next
        
End Function

Sub BorrarPJdeCuenta(ByVal name As String)
        
        On Error GoTo BorrarPJdeCuenta_Err
        

        Dim CantpjsNew  As Byte

        Dim CantpjsOld  As Byte

        Dim Indice      As Byte

        Dim pjs(1 To 8) As String

        Dim SiguientePJ As Byte

        Dim Cuenta      As String

100     Cuenta = ObtenerCuenta(name)
    
102     CantpjsOld = ObtenerCantidadDePersonajes(Cuenta)
    
        Dim i As Integer

104     For i = 1 To CantpjsOld
106         pjs(i) = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)

108         If pjs(i) = name Then
110             Indice = i
112             pjs(i) = ""

            End If

114     Next i
    
116     Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & Indice, "")
    
118     For i = 1 To CantpjsOld
120         pjs(i) = GetVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i)
122     Next i
    
124     For i = 1 To CantpjsOld

126         If pjs(i) = "" And i + 1 < 9 Then
128             pjs(i) = pjs(i + 1)
130             pjs(i + 1) = ""
132             Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i, pjs(i))
134             Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "PJ" & i + 1, "")

            End If

136     Next i
    
138     SiguientePJ = ObtenerCantidadDePersonajes(Cuenta) - 1
140     Call WriteVar(CuentasPath & UCase$(Cuenta) & ".act", "PERSONAJES", "Total", SiguientePJ)

        
        Exit Sub

BorrarPJdeCuenta_Err:
142     Call RegistrarError(Err.Number, Err.description, "ModCuentas.BorrarPJdeCuenta", Erl)
144     Resume Next
        
End Sub

Sub GrabarNuevoPjEnCuentaCharfile(ByVal UserCuenta As String, ByVal name As String)
        
        On Error GoTo GrabarNuevoPjEnCuentaCharfile_Err
        

        Dim cantidaddePersonajes As Byte

100     cantidaddePersonajes = ObtenerCantidadDePersonajes(UserCuenta)

102     Call WriteVar(CuentasPath & UCase$(UserCuenta) & ".act", "PERSONAJES", "Total", cantidaddePersonajes + 1)
104     Call WriteVar(CuentasPath & UCase$(UserCuenta) & ".act", "PERSONAJES", "PJ" & cantidaddePersonajes + 1, name)

        
        Exit Sub

GrabarNuevoPjEnCuentaCharfile_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GrabarNuevoPjEnCuentaCharfile", Erl)
108     Resume Next
        
End Sub

Sub BorrarCuenta(ByVal CuentaName As String)
        
        On Error GoTo BorrarCuenta_Err
        

100     If Database_Enabled Then
102         Call BorrarCuentaDatabase(CuentaName)
        Else

            'Primero borramos PJ POR PJ, se copia los personajes a la carpeta de personajes borrados
            Dim Cantpjs As Byte
        
104         Cantpjs = ObtenerCantidadDePersonajes(CuentaName)

            Dim Indice      As Byte

            Dim pjs(1 To 8) As String
        
            Dim i           As Integer

106         For i = 1 To Cantpjs
108             pjs(i) = GetVar(CuentasPath & UCase$(CuentaName) & ".act", "PERSONAJES", "PJ" & i)

110             If FileExist(CharPath & UCase$(pjs(i)) & ".chr", vbNormal) Then
112                 Call FileCopy(CharPath & UCase$(pjs(i)) & ".chr", DeletePath & UCase$(pjs(i)) & ".chr")
114                 Call Kill(CharPath & UCase$(pjs(i)) & ".chr")

                End If

116         Next i
    
118         If FileExist(CuentasPath & UCase$(CuentaName) & ".act", vbNormal) Then
120             Call FileCopy(CuentasPath & UCase$(CuentaName) & ".act", DeleteCuentasPath & UCase$(CuentaName))
122             Call Kill(CuentasPath & CuentaName & ".act")

            End If

        End If

        
        Exit Sub

BorrarCuenta_Err:
124     Call RegistrarError(Err.Number, Err.description, "ModCuentas.BorrarCuenta", Erl)
126     Resume Next
        
End Sub

Public Function ObtenerNivel(ByVal name As String) As Byte

        On Error GoTo ErrorHandler

100     ObtenerNivel = GetVar(CharPath & UCase$(name & ".chr"), "STATS", "ELV")

        Exit Function
ErrorHandler:
102     ObtenerNivel = 1

End Function

Public Function ObtenerCuerpo(ByVal name As String) As Integer

        On Error GoTo ErrorHandler

        Dim EstaMuerto As Byte

        Dim cuerpo     As Long

100     EstaMuerto = GetVar(CharPath & UCase$(name & ".chr"), "flags", "Muerto")

102     If EstaMuerto = 0 Then
104         cuerpo = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Body")
106         ObtenerCuerpo = cuerpo
        Else
108         ObtenerCuerpo = iCuerpoMuerto

        End If

        Exit Function
ErrorHandler:
110     ObtenerCuerpo = 1

End Function

Public Function ObtenerCabeza(ByVal name As String) As Integer

        On Error GoTo ErrorHandler

        Dim Head       As String

        Dim EstaMuerto As Byte

100     EstaMuerto = GetVar(CharPath & UCase$(name & ".chr"), "flags", "Muerto")

102     If EstaMuerto = 0 Then
104         Head = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Head")
        Else
106         Head = 0
        End If

108     ObtenerCabeza = Head

        Exit Function
ErrorHandler:
110     ObtenerCabeza = 1

End Function

Public Function ObtenerEscudo(ByVal name As String) As Integer

        On Error GoTo ErrorHandler

100     ObtenerEscudo = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Escudo")
        Exit Function
ErrorHandler:
102     ObtenerEscudo = 0

End Function

Public Function ObtenerArma(ByVal name As String) As Integer

        On Error GoTo ErrorHandler

100     ObtenerArma = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Arma")
        Exit Function
ErrorHandler:
102     ObtenerArma = 0

End Function

Public Function ObtenerCasco(ByVal name As String) As Integer

        On Error GoTo ErrorHandler

100     ObtenerCasco = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Casco")
        Exit Function
ErrorHandler:
102     ObtenerCasco = 0

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

112     If Not Database_Enabled Then
114         GetUserGuildIndex = GetUserGuildIndexCharfile(UserName)
        Else
116         GetUserGuildIndex = GetUserGuildIndexDatabase(UserName)

        End If

        
        Exit Function

GetUserGuildIndex_Err:
118     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GetUserGuildIndex", Erl)
120     Resume Next
        
End Function

Public Function GetUserGuildIndexCharfile(ByRef UserName As String) As Integer
        
        On Error GoTo GetUserGuildIndexCharfile_Err
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 26/09/2018
        '26/09/2018 CHOTS: Moved to FileIO
        '***************************************************
        Dim Temps As String
    
100     Temps = GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX")

102     If IsNumeric(Temps) Then
104         GetUserGuildIndexCharfile = CInt(Temps)
        Else
106         GetUserGuildIndexCharfile = 0

        End If

        
        Exit Function

GetUserGuildIndexCharfile_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GetUserGuildIndexCharfile", Erl)
110     Resume Next
        
End Function

Public Function GetUserGuildPedidosCharfile(ByVal UserName As String) As String
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo GetUserGuildPedidosCharfile_Err
        

100     GetUserGuildPedidosCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "Pedidos")

        
        Exit Function

GetUserGuildPedidosCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GetUserGuildPedidosCharfile", Erl)
104     Resume Next
        
End Function

Sub SaveUserGuildPedidosCharfile(ByVal UserName As String, ByVal Pedidos As String)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo SaveUserGuildPedidosCharfile_Err
        

100     Call WriteVar(CharPath & UserName & ".chr", "GUILD", "Pedidos", Pedidos)

        
        Exit Sub

SaveUserGuildPedidosCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveUserGuildPedidosCharfile", Erl)
104     Resume Next
        
End Sub

Sub SaveUserGuildMemberCharfile(ByVal UserName As String, ByVal guilds As String)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo SaveUserGuildMemberCharfile_Err
        

100     Call WriteVar(CharPath & UserName & ".chr", "GUILD", "Miembro", guilds)

        
        Exit Sub

SaveUserGuildMemberCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveUserGuildMemberCharfile", Erl)
104     Resume Next
        
End Sub

Sub SaveUserGuildIndexCharfile(ByVal UserName As String, ByVal GuildIndex As Integer)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo SaveUserGuildIndexCharfile_Err
        

100     Call WriteVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX", GuildIndex)

        
        Exit Sub

SaveUserGuildIndexCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveUserGuildIndexCharfile", Erl)
104     Resume Next
        
End Sub

Sub SaveUserGuildAspirantCharfile(ByVal UserName As String, ByVal AspirantIndex As Integer)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo SaveUserGuildAspirantCharfile_Err
        

100     Call WriteVar(CharPath & UserName & ".chr", "GUILD", "ASPIRANTEA", AspirantIndex)

        
        Exit Sub

SaveUserGuildAspirantCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveUserGuildAspirantCharfile", Erl)
104     Resume Next
        
End Sub

Sub SendCharacterInfoCharfile(ByVal UserIndex As Integer, ByVal UserName As String)
        
        On Error GoTo SendCharacterInfoCharfile_Err
        

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        Dim gName       As String

        Dim UserFile    As clsIniReader

        Dim Miembro     As String

        Dim GuildActual As Integer

        ' Get the character's current guild
100     GuildActual = GetUserGuildIndex(UserName)

102     If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
104         gName = "<" & GuildName(GuildActual) & ">"
        Else
106         gName = "Ninguno"

        End If
    
        'Get previous guilds
108     Miembro = GetUserGuildMember(UserName)

110     If Len(Miembro) > 400 Then
112         Miembro = ".." & Right$(Miembro, 400)

        End If

114     Set UserFile = New clsIniReader

116     With UserFile
118         .Initialize (CharPath & UserName & ".chr")
    
120         Call WriteCharacterInfo(UserIndex, UserName, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), .GetValue("STATS", "Banco"), .GetValue("GUILD", "Pedidos"), gName, Miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))

        End With
    
122     Set UserFile = Nothing

        
        Exit Sub

SendCharacterInfoCharfile_Err:
124     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SendCharacterInfoCharfile", Erl)
126     Resume Next
        
End Sub

Public Function GetUserGuildMemberCharfile(ByVal UserName As String) As String
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo GetUserGuildMemberCharfile_Err
        

100     GetUserGuildMemberCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "Miembro")

        
        Exit Function

GetUserGuildMemberCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GetUserGuildMemberCharfile", Erl)
104     Resume Next
        
End Function

Public Function GetUserGuildAspirantCharfile(ByVal UserName As String) As Integer
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo GetUserGuildAspirantCharfile_Err
        

100     GetUserGuildAspirantCharfile = val(GetVar(CharPath & UserName & ".chr", "GUILD", "ASPIRANTEA"))

        
        Exit Function

GetUserGuildAspirantCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GetUserGuildAspirantCharfile", Erl)
104     Resume Next
        
End Function

Public Function GetUserGuildRejectionReasonCharfile(ByVal UserName As String) As String
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo GetUserGuildRejectionReasonCharfile_Err
        

100     GetUserGuildRejectionReasonCharfile = GetVar(CharPath & UserName & ".chr", "GUILD", "MotivoRechazo")

        
        Exit Function

GetUserGuildRejectionReasonCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.GetUserGuildRejectionReasonCharfile", Erl)
104     Resume Next
        
End Function

Sub SaveUserGuildRejectionReasonCharfile(ByVal UserName As String, ByVal Reason As String)
        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 26/09/2018
        '***************************************************
        
        On Error GoTo SaveUserGuildRejectionReasonCharfile_Err
        

100     Call WriteVar(CharPath & UserName & ".chr", "GUILD", "MotivoRechazo", Reason)

        
        Exit Sub

SaveUserGuildRejectionReasonCharfile_Err:
102     Call RegistrarError(Err.Number, Err.description, "ModCuentas.SaveUserGuildRejectionReasonCharfile", Erl)
104     Resume Next
        
End Sub

Public Function ObtenerCriminal(ByVal name As String) As Byte

        On Error GoTo ErrorHandler
    
        Dim Criminal As Byte

100     If Database_Enabled Then
102         Criminal = GetUserStatusDatabase(name)
        Else
104         Criminal = GetVar(CharPath & UCase$(name & ".chr"), "FACCIONES", "Status")

        End If

106     If EsRolesMaster(name) Then
108         Criminal = 3
110     ElseIf EsConsejero(name) Then
112         Criminal = 4
114     ElseIf EsSemiDios(name) Then
116         Criminal = 5
118     ElseIf EsDios(name) Then
120         Criminal = 6
122     ElseIf EsAdmin(name) Then
124         Criminal = 7

        End If

126     ObtenerCriminal = Criminal

        Exit Function
ErrorHandler:
128     ObtenerCriminal = 1

End Function

Public Function ObtenerMapa(ByVal name As String) As String

        On Error GoTo ErrorHandler

        Dim Mapa As String

100     ObtenerMapa = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Position")
    
        Exit Function
ErrorHandler:
102     ObtenerMapa = "1-50-50"
    
End Function

Public Function ObtenerClase(ByVal name As String) As Byte

        On Error GoTo ErrorHandler

100     ObtenerClase = GetVar(CharPath & UCase$(name & ".chr"), "INIT", "Clase")

        Exit Function
ErrorHandler:
102     ObtenerClase = "1"

End Function
