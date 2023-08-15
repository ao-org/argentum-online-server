Attribute VB_Name = "Penas"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Public IP_Blacklist As New Dictionary

Public Sub CargarListaNegraUsuarios()
#If DEVELOPER = 0 Then
On Error GoTo CargarListaNegraUsuarios_Err
        Dim File   As clsIniManager
        Dim i      As Long
        Dim iKey   As String
        Dim iValue As String
100     If Not FileExist(DatPath & "Baneos.dat") Then Exit Sub
102     Set File = New clsIniManager
104     Call File.Initialize(DatPath & "Baneos.dat")
        Call IP_Blacklist.RemoveAll
108     For i = 0 To File.EntriesCount("IP") - 1
110        Call File.GetPair("IP", i, iKey, iValue)
            If Not IP_Blacklist.Exists(iKey) Then
112             Call IP_Blacklist.Add(iKey, iValue)
            End If
        Next

        Exit Sub
CargarListaNegraUsuarios_Err:
        Set File = Nothing
        Call TraceError(Err.Number, Err.Description, "Penas.CargarListaNegraUsuarios", Erl)
#End If
End Sub

Private Function GlobalChecks(ByVal BannerIndex As Integer, ByRef username As String) As Integer
        
        On Error GoTo GlobalChecks_Err

        Dim tUser As t_UserReference

100     GlobalChecks = False

102     If Not EsGM(BannerIndex) Then Exit Function

        ' Parseo los espacios en el Nick
104     If InStrB(UserName, "+") Then
106         UserName = Replace(UserName, "+", " ")
        End If

108     tUser = NameIndex(username)

110     If IsValidUserRef(tUser) Then

112         If tUser.ArrayIndex = BannerIndex Then
114             Call WriteConsoleMsg(BannerIndex, "No podes banearte a vos mismo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

            ' Estas tratando de banear a alguien con mas privilegios que vos, no va a pasar bro.
116         If CompararUserPrivilegios(tUser.ArrayIndex, BannerIndex) >= 0 Then
118             Call WriteConsoleMsg(BannerIndex, "No podes banear a al alguien de igual o mayor jerarquia.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If

        Else

120         If CompararPrivilegios(UserDarPrivilegioLevel(UserName), UserList(BannerIndex).flags.Privilegios) >= 0 Then
122             Call WriteConsoleMsg(BannerIndex, "No podes banear a al alguien de igual o mayor jerarquia.", e_FontTypeNames.FONTTYPE_INFO)
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
        
#If STRESSER = 1 Then
    Exit Sub
#End If
        
100     If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub

        ' Si no existe el personaje...
102     If Not PersonajeExiste(UserName) Then
104         Call WriteConsoleMsg(BannerIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If

106     If BANCheck(UserName) Then
108         Call WriteConsoleMsg(BannerIndex, "El usuario ya se encuentra baneado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
110     Call SaveBanDatabase(UserName, Razon, UserList(BannerIndex).Name)

        ' Registramos el baneo en los logs.
112     Call LogBanFromName(UserName, BannerIndex, Razon)

        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & UserList(BannerIndex).Name & " ha baneado a " & username & " debido a: " & LCase$(Razon) & ".", e_FontTypeNames.FONTTYPE_SERVER))

        ' Si estaba online, lo echamos.
116     Dim tUser As t_UserReference: tUser = NameIndex(username)
118     If IsValidUserRef(tUser) Then
            Call WriteDisconnect(tUser.ArrayIndex)
            Call CloseSocket(tUser.ArrayIndex)
        End If

        Exit Sub

BanPJ_Err:
120     Call TraceError(Err.Number, Err.Description, "Mod_Baneo.BanPJ")
122

End Sub

Public Sub BanPJWithoutGM(ByVal UserName As String, ByRef Razon As String)
        On Error GoTo BanPJWithoutGM_Err

        ' Si no existe el personaje...
102     If Not PersonajeExiste(UserName) Then
            Exit Sub
        End If

106     If BANCheck(UserName) Then
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
110     Call SaveBanDatabase(UserName, Razon, "el sistema")

100     Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", username, "BannedBy", "Ban automático (Posible BOT).")
        Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", UserName, "Reason", Razon)
        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » Ha baneado a " & username & " debido a: " & LCase$(Razon) & ".", e_FontTypeNames.FONTTYPE_SERVER))

        ' Si estaba online, lo echamos.
116     Dim tUser As t_UserReference: tUser = NameIndex(username)
118     If IsValidUserRef(tUser) Then
            Call WriteDisconnect(tUser.ArrayIndex)
            Call CloseSocket(tUser.ArrayIndex)
        End If

        Exit Sub

BanPJWithoutGM_Err:
120     Call TraceError(Err.Number, Err.Description, "Mod_Baneo.BanPJWithoutGM")
122

End Sub
Public Sub BanearCuenta(ByVal BannerIndex As Integer, ByVal UserName As String, ByVal Reason As String)
        On Error GoTo BanearCuenta_Err
        Dim CuentaID As Long
        
100     If Not GlobalChecks(BannerIndex, UserName) Then Exit Sub

        ' Obtenemos el ID de la cuenta
102     CuentaID = GetAccountIDDatabase(UserName)

        ' Me fijo que exista la cuenta.
104     If CuentaID <= 0 Then
106         Call WriteConsoleMsg(BannerIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
112     Call SaveBanCuentaDatabase(CuentaID, Reason, UserList(BannerIndex).Name)

        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & UserList(BannerIndex).Name & " ha baneado la cuenta de " & username & " debido a: " & Reason & ".", e_FontTypeNames.FONTTYPE_SERVER))

        ' Registramos el baneo en los logs.
116     Call LogGM(UserList(BannerIndex).Name, "Baneó la cuenta de " & username & " por: " & Reason)

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

Public Function DesbanearCuenta(ByVal BannerIndex As Integer, ByVal UserNameOEmail As String) As Boolean

        On Error GoTo DesbanearCuenta_Err
        
        ' Seteamos is_banned = 0 en la DB
        If InStr(1, UserNameOEmail, "@") Then
            DesbanearCuenta = Execute("UPDATE account SET is_banned = false WHERE email = ?", UserNameOEmail)
        Else
            DesbanearCuenta = Execute("UPDATE `account` INNER JOIN `user` ON user.account_id=account.id SET account.is_banned=FALSE WHERE user.name = ?", UserNameOEmail)
        End If

        Exit Function

DesbanearCuenta_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.DesbanearCuenta", Erl)
End Function

Public Sub BanearIP(ByVal BannerIndex As Integer, ByVal UserName As String, ByVal IP As String, Optional ByVal Email As String)
        On Error GoTo BanearIP_Err
        
#If STRESSER = 1 Then
    Exit Sub
#End If
        ' Lo guardo en Baneos.dat
100     Call WriteVar(DatPath & "Baneos.dat", "IP", IP, UserName)

        If LenB(UserName) > 0 Then
            If Not (val(mid(UserName, 1, 1)) > 0) Then
                Call Execute("UPDATE account set is_banned = true where UPPER(email) = ?;", UCase$(Email))
                Call BanPJWithoutGM(UserName, "Por ban IP.")
            End If
        End If
        
        ' Lo guardo en memoria.
102     Call IP_Blacklist.Add(IP, UserName)

        ' Registramos el des-baneo en los logs.
104     Call LogGM(UserList(BannerIndex).Name, "Baneó la IP: " & IP & " de " & username)

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

        ' Registramos el des-baneo en los logs.

        Exit Sub

DesbanearIP_Err:
        Call TraceError(Err.Number, Err.Description, "Penas.DesbanearIP", Erl)
End Sub
