Attribute VB_Name = "Penas"
Option Explicit

Public Enum e_LoadBlacklistFlags
    LoadIPs = &H1   ' 2 ^ 0
    LoadHDs = &H2   ' 2 ^ 1
    LoadMACs = &H4  ' 2 ^ 2
    LoadAll = LoadIPs Or LoadHDs Or LoadMACs
End Enum

Public IP_Blacklist As New Dictionary
Public HD_Blacklist As New Dictionary
Public MAC_Blacklist As New Dictionary

Public Sub CargarListaNegraUsuarios(ByVal LoadFlags As e_LoadBlacklistFlags)
        
        On Error GoTo CargarListaNegraUsuarios_Err

        Dim File   As clsIniManager
        Dim i      As Long
        Dim iKey   As String
        Dim iValue As String

100     If Not FileExist(DatPath & "Baneos.dat") Then Exit Sub

102     Set File = New clsIniManager
104     Call File.Initialize(DatPath & "Baneos.dat")

106     If (LoadFlags And LoadIPs) Then

            ' IP's
108         For i = 0 To File.EntriesCount("IP") - 1
110             Call File.GetPair("IP", i, iKey, iValue)
112             Call IP_Blacklist.Add(iKey, iValue)
            Next

        End If

114     If (LoadFlags And LoadHDs) Then

            ' HD's
116         For i = 0 To File.EntriesCount("HD") - 1
118             Call File.GetPair("HD", i, iKey, iValue)
120             Call HD_Blacklist.Add(iKey, iValue)
            Next

        End If

122     If (LoadFlags And LoadMACs) Then

            ' MAC's
124         For i = 0 To File.EntriesCount("MAC") - 1
126             Call File.GetPair("MAC", i, iKey, iValue)
128             Call MAC_Blacklist.Add(iKey, iValue)
            Next

        End If
    
        Exit Sub

CargarListaNegraUsuarios_Err:
        Set File = Nothing
        Call TraceError(Err.Number, Err.Description, "Penas.CargarListaNegraUsuarios", Erl)

        
        
End Sub

Private Function GlobalChecks(ByVal BannerIndex, ByRef UserName As String) As Integer
        
        On Error GoTo GlobalChecks_Err

        Dim TargetIndex As Integer

100     GlobalChecks = False

102     If Not EsGM(BannerIndex) Then Exit Function

        ' Parseo los espacios en el Nick
104     If InStrB(UserName, "+") Then
106         UserName = Replace(UserName, "+", " ")
        End If

108     TargetIndex = NameIndex(UserName)

110     If TargetIndex Then

112         If TargetIndex = BannerIndex Then
114             Call WriteConsoleMsg(BannerIndex, "No podes banearte a vos mismo.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

            ' Estas tratando de banear a alguien con mas privilegios que vos, no va a pasar bro.
116         If CompararUserPrivilegios(TargetIndex, BannerIndex) >= 0 Then
118             Call WriteConsoleMsg(BannerIndex, "No podes banear a al alguien de igual o mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

        Else

120         If CompararPrivilegios(UserDarPrivilegioLevel(UserName), UserList(BannerIndex).flags.Privilegios) >= 0 Then
122             Call WriteConsoleMsg(BannerIndex, "No podes banear a al alguien de igual o mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

        End If

        ' Se llegó hasta acá, todo bien!
124     GlobalChecks = True

        
        Exit Function

GlobalChecks_Err:
    Call TraceError(Err.Number, Err.Description, "Penas.GlobalChecks", Erl)
    
    
End Function

Public Sub BanPJ(ByVal BannerIndex As Integer, ByVal UserName As String, ByRef Razon As String)
        On Error GoTo BanPJ_Err

100     If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub

        ' Si no existe el personaje...
102     If Not PersonajeExiste(UserName) Then
104         Call WriteConsoleMsg(BannerIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If

106     If BANCheck(UserName) Then
108         Call WriteConsoleMsg(BannerIndex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
110     Call SaveBanDatabase(UserName, Razon, UserList(BannerIndex).Name)

        ' Registramos el baneo en los logs.
112     Call LogBanFromName(UserName, BannerIndex, Razon)

        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & UserList(BannerIndex).Name & " ha baneado a " & UserName & " debido a: " & LCase$(Razon) & ".", FontTypeNames.FONTTYPE_SERVER))

        ' Si estaba online, lo echamos.
116     Dim tUser As Integer: tUser = NameIndex(UserName)
118     If tUser > 0 Then
            Call WriteDisconnect(tUser)
            Call CloseSocket(tUser)
        End If

        Exit Sub

BanPJ_Err:
120     Call TraceError(Err.Number, Err.Description, "Mod_Baneo.BanPJ")
122

End Sub

Public Sub BanearCuenta(ByVal BannerIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal Reason As String)
        
        On Error GoTo BanearCuenta_Err
        

        Dim CuentaID As Integer

100     If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub

        ' Obtenemos el ID de la cuenta
102     CuentaID = GetAccountIDDatabase(UserName)

        ' Me fijo que exista la cuenta.
104     If CuentaID <= 0 Then
106         Call WriteConsoleMsg(BannerIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub

        End If

108     If ObtenerBaneo(UserName) Then
110         Call WriteConsoleMsg(BannerIndex, "La cuenta ya se encuentra baneada.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
112     Call SaveBanCuentaDatabase(CuentaID, Reason, UserList(BannerIndex).Name)

        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & UserList(BannerIndex).Name & " ha baneado la cuenta de " & UserName & " debido a: " & Reason & ".", FontTypeNames.FONTTYPE_SERVER))

        ' Registramos el baneo en los logs.
116     Call LogGM(UserList(BannerIndex).Name, "Baneó la cuenta de " & UserName & " por: " & Reason)

        ' Echo a todos los logueados en esta cuenta
        Dim i As Long
118     For i = 1 To LastUser

120         If UserList(i).AccountID = CuentaID Then
122             Call WriteShowMessageBox(i, "Has sido baneado del servidor. Motivo: " & Reason)
                Call WriteDisconnect(i)
124             Call CloseSocket(i)

            End If

        Next

        Exit Sub

BanearCuenta_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.BanearCuenta", Erl)
        
        
End Sub

Public Function DesbanearCuenta(ByVal BannerIndex As Integer, ByVal AccountID As Long) As Boolean

        On Error GoTo DesbanearCuenta_Err

102     If Not GetBaneoAccountId(AccountID) Then
104         Call WriteConsoleMsg(BannerIndex, "La cuenta no se encuentra baneada.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

        ' Seteamos is_banned = 0 en la DB
110     Call SetDBValue("account", "is_banned", 0, "id", AccountID)

        DesbanearCuenta = True

        Exit Function

DesbanearCuenta_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.DesbanearCuenta", Erl)
        
        
End Function

Public Sub BanearIP(ByVal BannerIndex As Integer, ByVal UserName As String, ByVal IP As String)
        
        
        On Error GoTo BanearIP_Err
        
        ' Lo guardo en Baneos.dat
100     Call WriteVar(DatPath & "Baneos.dat", "IP", IP, UserName)

        ' Lo guardo en memoria.
102     Call IP_Blacklist.Add(IP, UserName)

        ' Agregar a la regla de firewall
        'Dim i As Long
        'Dim NewIPs As String
        'For i = 0 To IP_Blacklist.Count - 1
        '    NewIPs = NewIPs & IP_Blacklist(i) & ","
        'Next

        'Call Shell("netsh.exe advfirewall firewall set rule name=""Lista Negra IPs"" dir=in remoteip=" & NewIPs) ' Turbio esto

        ' Registramos el des-baneo en los logs.
104     Call LogGM(UserList(BannerIndex).Name, "Baneó la IP: " & IP & " de " & UserName)

        Exit Sub

BanearIP_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.BanearIP", Erl)
        
        
End Sub

Public Sub DesbanearIP(ByVal IP As String, ByVal UnbannerIndex As Integer)
        
        On Error GoTo DesbanearIP_Err

        ' Lo saco de la memoria.
100     If IP_Blacklist.Exists(IP) Then Call IP_Blacklist.Remove(IP)

        ' Lo saco del archivo.
102     Call WriteVar(DatPath & "Baneos.dat", "IP", IP, vbNullString)
        
        ' Modificar en la regla de firewall
        'Dim i As Long
        'Dim NewIPs As String
        'For i = 0 To IP_Blacklist.Count - 1
        '    ' Meto todas MENOS la que vamos a desbanear
        '    If IP_Blacklist(i) <> ip Then
        '        NewIPs = NewIPs & IP_Blacklist(i) & ","
        '    End If
        'Next
        
        'Call Shell("netsh.exe advfirewall firewall set rule name=""Lista IPs Prohibidas"" dir=in remoteip=" & NewIPs)
        
        ' Registramos el des-baneo en los logs.
104     Call LogGM(UserList(UnbannerIndex).Name, "Des-Baneó la IP: " & IP & " de " & IP_Blacklist(IP))

        Exit Sub

DesbanearIP_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.DesbanearIP", Erl)
        
        
End Sub

Public Sub BanearHDMAC(ByVal BannerIndex As Integer, ByVal UserName As String)
        
        On Error GoTo BanearHDMAC_Err

        Dim Cuenta As String
        Dim HDSerial As String
        Dim MacAddress As String

100     If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub

102     Cuenta = ObtenerCuenta(UserName)

104     If LenB(Cuenta) = 0 Then
106         Call WriteConsoleMsg(BannerIndex, "La cuenta no existe.", FontTypeNames.FONTTYPE_TALK)
            Exit Sub

        End If

108     HDSerial = ObtenerHDserial(Cuenta)
110     MacAddress = ObtenerMacAdress(Cuenta)

        ' Lo guardo en memoria.
112     Call HD_Blacklist.Add(HDSerial, UserName)
114     Call MAC_Blacklist.Add(MacAddress, UserName)

        ' Lo guardo en Baneos.dat
116     Call WriteVar(DatPath & "Baneos.dat", "HD", HDSerial, UserName)
118     Call WriteVar(DatPath & "Baneos.dat", "MAC", MacAddress, UserName)
    
        ' Lo kickeo
120     Dim TargetIndex As Integer: TargetIndex = NameIndex(UserName)
122     If TargetIndex > 0 Then Call CloseSocket(TargetIndex)
    
        ' Registramos el baneo en los logs.
124     Call LogGM(UserList(BannerIndex).Name, "Aplicó Tolerancia 0 a: " & UserName & " con Serial HD: " & HDSerial & " y MAC Address: " & MacAddress)

        Exit Sub

BanearHDMAC_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.BanearHDMAC", Erl)
        

End Sub

Public Sub DesbanearHDMAC(ByVal UserName As String)

        On Error GoTo DesbanearHDMAC_Err
        

        Dim Cuenta As String
        Dim HDSerial As String
        Dim MacAddress As String

100     Cuenta = ObtenerCuenta(UserName)

102     If LenB(Cuenta) = 0 Then Exit Sub

104     HDSerial = ObtenerHDserial(Cuenta)
106     MacAddress = ObtenerMacAdress(Cuenta)

        ' Lo guardo en memoria.
108     Call HD_Blacklist.Remove(HDSerial)
110     Call MAC_Blacklist.Remove(MacAddress)

        ' Lo guardo en Baneos.dat
112     Call WriteVar(DatPath & "Baneos.dat", "HD", HDSerial, vbNullString)
114     Call WriteVar(DatPath & "Baneos.dat", "MAC", MacAddress, vbNullString)

        ' Registramos el baneo en los logs.
116     Call LogDesarrollo("Le quitó la Tolerancia 0 a: " & UserName & " con Serial HD: " & HDSerial & " y MAC Address: " & MacAddress)

        Exit Sub

DesbanearHDMAC_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.DesbanearHDMAC", Erl)
        

End Sub

