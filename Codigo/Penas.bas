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

Private Function GlobalChecks(ByVal BannerIndex As Integer, ByRef username As String) As Integer
    On Error GoTo GlobalChecks_Err
    Dim tUser As t_UserReference
    GlobalChecks = False
    If Not EsGM(BannerIndex) Then Exit Function
    ' Parseo los espacios en el Nick
    If InStrB(username, "+") Then
        username = Replace(username, "+", " ")
    End If
    tUser = NameIndex(username)
    If IsValidUserRef(tUser) Then
        If tUser.ArrayIndex = BannerIndex Then
            Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1841, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1841=No podés banearte a vos mismo.
            Exit Function
        End If
        ' Estas tratando de banear a alguien con mas privilegios que vos, no va a pasar bro.
        If CompararUserPrivilegios(tUser.ArrayIndex, BannerIndex) >= 0 Then
            Call WriteLocaleMsg(BannerIndex, 2069, e_FontTypeNames.FONTTYPE_INFO) ' Msg2069="No podes banear a al alguien de igual o mayor jerarquia."
            Exit Function
        End If
    Else
        If CompararPrivilegios(UserDarPrivilegioLevel(username), UserList(BannerIndex).flags.Privilegios) >= 0 Then
            Call WriteLocaleMsg(BannerIndex, 2070, e_FontTypeNames.FONTTYPE_INFO) ' Msg2070="No podes banear a al alguien de igual o mayor jerarquia."
            Exit Function
        End If
    End If
    ' Se llegó hasta acá, todo bien!
    GlobalChecks = True
    Exit Function
GlobalChecks_Err:
    Call TraceError(Err.Number, Err.Description, "Penas.GlobalChecks", Erl)
End Function

Public Sub BanPJ(ByVal BannerIndex As Integer, ByVal username As String, ByRef Razon As String)
    On Error GoTo BanPJ_Err
    If Not GlobalChecks(BannerIndex, username) Then Exit Sub
    ' Si no existe el personaje...
    If Not PersonajeExiste(username) Then
        Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1842, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1842=El personaje no existe.
        Exit Sub
    End If
    If BANCheck(username) Then
        Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1843, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1843=El usuario ya se encuentra baneado.
        Exit Sub
    End If
    ' Guardamos el estado de baneado en la base de datos.
    Call SaveBanDatabase(username, Razon, UserList(BannerIndex).name)
    ' Registramos el baneo en los logs.
    Call LogBanFromName(username, BannerIndex, Razon)
    ' Le buchoneamos al mundo.
    Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1699, UserList(BannerIndex).name & "¬" & username & "¬" & LCase$(Razon), _
            e_FontTypeNames.FONTTYPE_SERVER))  'Msg1699=Servidor » ¬1 ha baneado a ¬2 debido a: ¬3.
    ' Si estaba online, lo echamos.
    Dim tUser As t_UserReference: tUser = NameIndex(username)
    If IsValidUserRef(tUser) Then
        Call WriteDisconnect(tUser.ArrayIndex)
        Call CloseSocket(tUser.ArrayIndex)
    End If
    Exit Sub
BanPJ_Err:
    Call TraceError(Err.Number, Err.Description, "Mod_Baneo.BanPJ")
End Sub

Public Sub BanPJWithoutGM(ByVal username As String, ByRef Razon As String)
    On Error GoTo BanPJWithoutGM_Err
    ' Si no existe el personaje...
    If Not PersonajeExiste(username) Then
        Exit Sub
    End If
    If BANCheck(username) Then
        Exit Sub
    End If
    ' Guardamos el estado de baneado en la base de datos.
    Call SaveBanDatabase(username, Razon, "el sistema")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", username, "BannedBy", "Ban automático (Posible BOT).")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", username, "Reason", Razon)
    ' Le buchoneamos al mundo.
    Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1700, username & "¬" & LCase$(Razon), e_FontTypeNames.FONTTYPE_SERVER))  'Msg1700=Servidor » Ha baneado a ¬1 debido a: ¬2.
    ' Si estaba online, lo echamos.
    Dim tUser As t_UserReference: tUser = NameIndex(username)
    If IsValidUserRef(tUser) Then
        Call WriteDisconnect(tUser.ArrayIndex)
        Call CloseSocket(tUser.ArrayIndex)
    End If
    Exit Sub
BanPJWithoutGM_Err:
    Call TraceError(Err.Number, Err.Description, "Mod_Baneo.BanPJWithoutGM")
End Sub

Public Sub BanearCuenta(ByVal BannerIndex As Integer, ByVal username As String, ByVal Reason As String)
    On Error GoTo BanearCuenta_Err
    Dim CuentaID As Long
    If Not GlobalChecks(BannerIndex, username) Then Exit Sub
    ' Obtenemos el ID de la cuenta
    CuentaID = GetAccountIDDatabase(username)
    ' Me fijo que exista la cuenta.
    If CuentaID <= 0 Then
        Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1842, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1842=El personaje no existe.
        Exit Sub
    End If
    ' Guardamos el estado de baneado en la base de datos.
    Call SaveBanCuentaDatabase(CuentaID, Reason, UserList(BannerIndex).name)
    ' Le buchoneamos al mundo.
    Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1701, UserList(BannerIndex).name & "¬" & username & "¬" & Reason, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1701=Servidor » ¬1 ha baneado la cuenta de ¬2 debido a: ¬3.
    ' Registramos el baneo en los logs.
    Call LogGM(UserList(BannerIndex).name, "Baneó la cuenta de " & username & " por: " & Reason)
    ' Echo a todos los logueados en esta cuenta
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).AccountID = CuentaID Then
            Call WriteShowMessageBox(i, 1785, Reason) 'Msg1785=Has sido baneado del servidor. Motivo: ¬1
            Call WriteDisconnect(i)
            Call CloseSocket(i)
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
