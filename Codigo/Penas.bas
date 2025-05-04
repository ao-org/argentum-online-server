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

100     GlobalChecks = False

102     If Not EsGM(BannerIndex) Then Exit Function

        ' Parseo los espacios en el Nick
104     If InStrB(UserName, "+") Then
106         UserName = Replace(UserName, "+", " ")
        End If

108     tUser = NameIndex(username)

110     If IsValidUserRef(tUser) Then

112         If tUser.ArrayIndex = BannerIndex Then
114             Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1841, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1841=No podés banearte a vos mismo.
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
104         Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1842, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1842=El personaje no existe.
            Exit Sub
        End If

106     If BANCheck(UserName) Then
108         Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1843, vbNullString, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1843=El usuario ya se encuentra baneado.
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
110     Call SaveBanDatabase(UserName, Razon, UserList(BannerIndex).Name)

        ' Registramos el baneo en los logs.
112     Call LogBanFromName(UserName, BannerIndex, Razon)

        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1699, UserList(BannerIndex).name & "¬" & username & "¬" & LCase$(Razon), e_FontTypeNames.FONTTYPE_SERVER))  'Msg1699=Servidor » ¬1 ha baneado a ¬2 debido a: ¬3.

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
114     Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1700, username & "¬" & LCase$(Razon), e_FontTypeNames.FONTTYPE_SERVER))  'Msg1700=Servidor » Ha baneado a ¬1 debido a: ¬2.


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
106         Call WriteConsoleMsg(BannerIndex, PrepareMessageLocaleMsg(1842, vbNullString, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1842=El personaje no existe.
            Exit Sub
        End If

        ' Guardamos el estado de baneado en la base de datos.
112     Call SaveBanCuentaDatabase(CuentaID, Reason, UserList(BannerIndex).Name)

        ' Le buchoneamos al mundo.
114     Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1701, UserList(BannerIndex).name & "¬" & username & "¬" & Reason, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1701=Servidor » ¬1 ha baneado la cuenta de ¬2 debido a: ¬3.

        ' Registramos el baneo en los logs.
116     Call LogGM(UserList(BannerIndex).name, "Baneó la cuenta de " & username & " por: " & Reason)

        ' Echo a todos los logueados en esta cuenta
        Dim i As Long
118     For i = 1 To LastUser

120         If UserList(i).AccountID = CuentaID Then
122             Call WriteShowMessageBox(i, 1785, Reason) 'Msg1785=Has sido baneado del servidor. Motivo: ¬1
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


