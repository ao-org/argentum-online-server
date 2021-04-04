Attribute VB_Name = "mod_Baneos"
Option Explicit

Public Enum e_LoadBlacklistFlags
    LoadAll = &H1   ' 2 ^ 0
    LoadIPs = &H2   ' 2 ^ 1
    LoadHDs = &H4   ' 2 ^ 2
    LoadMACs = &H8  ' 2 ^ 3
End Enum

Public IP_Blacklist As New Dictionary
Public HD_Blacklist As New Dictionary
Public MAC_Blacklist As New Dictionary

Public Sub CargarListaNegraUsuarios(ByVal LoadFlags As e_LoadBlacklistFlags)
    
    Dim File As clsIniManager
    Dim i As Long
    Dim iKey As String
    Dim iValue As String
    
    Set File = New clsIniManager
    Call File.Initialize(DatPath & "Baneos.ini")
    
    If (LoadFlags And (LoadAll Or LoadIPs)) Then
    
        ' IP's
        For i = 0 To File.EntriesCount("IP") - 1
            Call File.GetPair("IP", i, iKey, iValue)
            Call IP_Blacklist.Add(iKey, iValue)
        Next
        
    End If
    
    If (LoadFlags And (LoadAll Or LoadHDs)) Then
    
        ' HD's
        For i = 0 To File.EntriesCount("HD") - 1
            Call File.GetPair("HD", i, iKey, iValue)
            Call HD_Blacklist.Add(iKey, iValue)
        Next
        
    End If
    
    If (LoadFlags And (LoadAll Or LoadMACs)) Then
    
        ' MAC's
        For i = 0 To File.EntriesCount("MAC") - 1
            Call File.GetPair("MAC", i, iKey, iValue)
            Call MAC_Blacklist.Add(iKey, iValue)
        Next
        
    End If
    
End Sub

Private Function GlobalChecks(ByVal BannerIndex, ByRef UserName As String) As Integer
    
    Dim TargetIndex As Integer
    
    GlobalChecks = False
    
    If Not EsGM(BannerIndex) Then Exit Function

    ' Parseo los espacios en el Nick
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    TargetIndex = NameIndex(UserName)
    
    If TargetIndex Then
        
        If TargetIndex = BannerIndex Then
            Call WriteConsoleMsg(BannerIndex, "No podes banearte a vos mismo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Estas tratando de banear a alguien con mas privilegios que vos, no va a pasar bro.
        If CompararUserPrivilegios(TargetIndex, BannerIndex) >= 0 Then
            Call WriteConsoleMsg(BannerIndex, "No podes banear a al alguien de igual o mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
    Else
    
        If CompararPrivilegios(UserDarPrivilegioLevel(UserName), UserList(BannerIndex).flags.Privilegios) >= 0 Then
            Call WriteConsoleMsg(BannerIndex, "No podes banear a al alguien de igual o mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    
    End If
    
    ' Se llegó hasta acá, todo bien!
    GlobalChecks = True
    
End Function

Public Sub BanPJ(ByVal BannerIndex As Integer, ByVal UserName As String, ByRef Razon As String)

    If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub
    
    ' Busco el UserIndex del PJ
    Dim tUser As Integer: tUser = NameIndex(UserName)
    
    ' Si no existe el personaje...
    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(BannerIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    End If
    
    If BANCheck(UserName) Then
        Call WriteConsoleMsg(BannerIndex, "El usuario ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Registramos el baneo en los logs.
    Call LogBanFromName(UserName, BannerIndex, Razon)
    
    ' Le buchoneamos al mundo.
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & UserList(BannerIndex).Name & " ha baneado a " & UserName & " debido a: " & LCase$(Razon) & ".", FontTypeNames.FONTTYPE_SERVER))
    
    ' Si estaba online, lo echamos.
    If tUser > 0 Then Call CloseSocket(tUser)
    
    ' Guardamos el estado de baneado en la base de datos.
    Call SaveBanDatabase(UserName, Razon, UserList(BannerIndex).Name)
    
End Sub

Public Sub BanearCuenta(ByVal BannerIndex As Integer, _
                        ByVal UserName As String, _
                        ByVal Reason As String)
    
    Dim CuentaID As Integer
    
    If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub
    
    ' Obtenemos el ID de la cuenta
    CuentaID = GetAccountIDDatabase(UserName)
    
    ' Me fijo que exista la cuenta.
    If CuentaID <= 0 Then
        Call WriteConsoleMsg(BannerIndex, "El personaje no existe.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
    
    End If

    If ObtenerBaneo(UserName) Then
        Call WriteConsoleMsg(BannerIndex, "La cuenta ya se encuentra baneada.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Guardamos el estado de baneado en la base de datos.
    Call SaveBanCuentaDatabase(CuentaID, Reason, UserList(BannerIndex).Name)
    
    ' Le buchoneamos al mundo.
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & UserList(BannerIndex).Name & " ha baneado la cuenta de " & UserName & " debido a: " & Reason & ".", FontTypeNames.FONTTYPE_SERVER))
    
    ' Registramos el baneo en los logs.
    Call LogGM(UserList(BannerIndex).Name, "Baneó la cuenta de " & UserName & " por: " & Reason)
    
    ' Echo a todos los logueados en esta cuenta
    Dim i As Long
    For i = 1 To LastUser

        If UserList(i).AccountId = CuentaID Then
            Call WriteShowMessageBox(i, "Has sido baneado del servidor. Motivo: " & Reason)
            Call CloseSocket(i)

        End If

    Next

End Sub

Public Sub DesbanearCuenta(ByVal BannerIndex As Integer, ByVal UserName As String)

    If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub
    
    If Not ObtenerBaneo(UserName) Then
        Call WriteConsoleMsg(BannerIndex, "La cuenta no se encuentra baneada.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Busco el ID de la cuenta baneada a partir del nick de uno de sus PJ's
    Call MakeQuery("SELECT `account_id`, `account`.email FROM `user` INNER JOIN `account` ON `user`.account_id = account.id WHERE `account`.is_banned = TRUE AND UPPER(`user`.name) = ?;", False, UCase$(UserName))
    
    ' Encontre algo?
    If QueryData Is Nothing Then Exit Sub
    
    ' Seteamos is_banned = 0 en la DB
    Call SetDBValue("account", "is_banned", 0, "id", QueryData!account_id)
    
    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & UserList(BannerIndex).Name & " ha desbaneado la cuenta de " & UserName & "(" & QueryData!email & ").", FontTypeNames.FONTTYPE_SERVER))
    
End Sub

Public Sub BanearIP(ByVal BannerIndex As Integer, ByVal UserName As String, ByVal IP As String)

    ' Lo guardo en Baneos.dat
    Call WriteVar(DatPath & "Baneos.dat", "IP", UserName, IP)

    ' Lo guardo en memoria.
    Call IP_Blacklist.Add(IP, UserName)
    
    ' TODO: Agregar regla de firewall
    
    ' Registramos el des-baneo en los logs.
    Call LogGM(UserList(BannerIndex).Name, "Baneó la IP: " & IP & " de " & UserName)
    
End Sub

Public Sub DesbanearIP(ByVal IP As String, ByVal UnbannerIndex As Integer)

    ' Lo saco de la memoria.
    If IP_Blacklist.Exists(IP) Then Call IP_Blacklist.Remove(IP)
        
    ' Lo saco del archivo.
    Call WriteVar(DatPath & "Baneos.dat", "IP", GetVar(DatPath & "Baneos.dat", "IP", IP), vbNullString)
    
    ' Registramos el des-baneo en los logs.
    Call LogGM(UserList(UnbannerIndex).Name, "Des-Baneó la IP: " & IP & " de " & IP_Blacklist(IP))
    
End Sub

Public Sub BanearHDMAC(ByVal BannerIndex As Integer, ByVal UserName As String)
    
    Dim Cuenta As String
    Dim HDSerial As String
    Dim MacAddress As String
    
    If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub
    
    Cuenta = ObtenerCuenta(UserName)
    
    If LenB(Cuenta) = 0 Then
        Call WriteConsoleMsg(BannerIndex, "La cuenta no existe.", FontTypeNames.FONTTYPE_TALK)
        Exit Sub
        
    End If
    
    HDSerial = ObtenerHDserial(Cuenta)
    MacAddress = ObtenerMacAdress(Cuenta)
    
    ' Lo guardo en memoria.
    Call HD_Blacklist.Add(HDSerial, UserName)
    Call MAC_Blacklist.Add(MacAddress, UserName)
    
    ' Lo guardo en Baneos.dat
    Call WriteVar(DatPath & "Baneos.dat", "HD", HDSerial, UserName)
    Call WriteVar(DatPath & "Baneos.dat", "MAC", MacAddress, UserName)
    
    ' Registramos el baneo en los logs.
    Call LogGM(UserList(BannerIndex).Name, "Aplicó Tolerancia 0 a: " & UserName & " con Serial HD: " & HDSerial & " y MAC Address: " & MacAddress)
    
End Sub

Public Sub DesbanearHDMAC(ByVal UserName As String)
    
    Dim Cuenta As String
    Dim HDSerial As String
    Dim MacAddress As String
    
    Cuenta = ObtenerCuenta(UserName)
    
    If LenB(Cuenta) = 0 Then Exit Sub

    HDSerial = ObtenerHDserial(Cuenta)
    MacAddress = ObtenerMacAdress(Cuenta)
    
    ' Lo guardo en memoria.
    Call HD_Blacklist.Remove(HDSerial)
    Call MAC_Blacklist.Remove(MacAddress)
    
    ' Lo guardo en Baneos.dat
    Call WriteVar(DatPath & "Baneos.dat", "HD", HDSerial, vbNullString)
    Call WriteVar(DatPath & "Baneos.dat", "MAC", MacAddress, vbNullString)
    
    ' Registramos el baneo en los logs.
    Call LogDesarrollo("Le quitó la Tolerancia 0 a: " & UserName & " con Serial HD: " & HDSerial & " y MAC Address: " & MacAddress)
    
End Sub

