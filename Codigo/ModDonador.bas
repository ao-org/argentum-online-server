Attribute VB_Name = "ModDonador"

Public Donadores        As New Collection

Public Codigo(1 To 800) As CodigoData

Public NumeroCodigos    As Integer

Public Type CodigoData

    Key As String
    Tipo As Byte
    Cantidad As Integer
    Usado As Byte

End Type

Option Explicit

Public Sub DonadorTiempo(ByVal nombre As String, ByVal dias As Integer)
        
        On Error GoTo DonadorTiempo_Err
        

100     If DonadorCheck(nombre) = 0 Then

            Dim tDon As TDonador
    
102         Set tDon = New TDonador
    
104         tDon.name = nombre
106         tDon.FechaExpiracion = (Now + dias)
    
108         Call Donadores.Add(tDon)
110         Call SaveDonador(Donadores.Count)
112         Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & nombre & " agrego " & dias & " días de donador.", FontTypeNames.FONTTYPE_New_DONADOR))
114         Call LogearEventoDeDonador("Se agregaron " & dias & " a la cuenta " & nombre & ".")
        Else

            'UserList(UserIndex).donador.FechaExpiracion = UserList(UserIndex).donador.FechaExpiracion + dias
    
            Dim LoopC As Integer

116         For LoopC = 1 To Donadores.Count

118             If UCase$(Donadores(LoopC).name) = UCase$(nombre) Then
120                 Donadores(LoopC).FechaExpiracion = Donadores(LoopC).FechaExpiracion + dias
122                 Call SaveDonador(LoopC)
                    Exit For

                End If

124         Next LoopC

        End If

        
        Exit Sub

DonadorTiempo_Err:
126     Call RegistrarError(Err.Number, Err.description, "ModDonador.DonadorTiempo", Erl)
128     Resume Next
        
End Sub

Sub SaveDonadores()
        
        On Error GoTo SaveDonadores_Err
        

        Dim num As Integer

100     Call WriteVar(DatPath & "Donadores.dat", "INIT", "NumeroDonadores", val(Donadores.Count))

102     For num = 1 To Donadores.Count
104         Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "USER", Donadores(num).name)
106         Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "FECHAEXPIRACION", Donadores(num).FechaExpiracion)
        Next

        
        Exit Sub

SaveDonadores_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModDonador.SaveDonadores", Erl)
110     Resume Next
        
End Sub

Sub SaveDonador(num As Integer)
        
        On Error GoTo SaveDonador_Err
        

100     Call WriteVar(DatPath & "Donadores.dat", "INIT", "NumeroDonadores", Donadores.Count)
102     Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "USER", Donadores(num).name)
104     Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "FECHAEXPIRACION", Donadores(num).FechaExpiracion)

106     Call WriteVar(CuentasPath & Donadores(num).name & ".act", "DONADOR", "DONADOR", "1")
108     Call WriteVar(CuentasPath & Donadores(num).name & ".act", "DONADOR", "FECHAEXPIRACION", Donadores(num).FechaExpiracion)

        
        Exit Sub

SaveDonador_Err:
110     Call RegistrarError(Err.Number, Err.description, "ModDonador.SaveDonador", Erl)
112     Resume Next
        
End Sub

Sub AgregarCreditosDonador(name As String, Cantidad As Long)
        
        On Error GoTo AgregarCreditosDonador_Err
        

        Dim creditos As Long

100     creditos = CreditosDonadorCheck(name) + Cantidad

        'Call LogearEventoDeDonador("Se agregaron " & Cantidad & " creditos a la cuenta " & Name & ".")
102     Call WriteVar(CuentasPath & UCase$(name & ".act"), "DONADOR", "CREDITOS", creditos)

        'Call AgregarCompra(Name, Date & " - Se agregaron " & Cantidad & " creditos a la cuenta " & Name & ".")
        
        Exit Sub

AgregarCreditosDonador_Err:
104     Call RegistrarError(Err.Number, Err.description, "ModDonador.AgregarCreditosDonador", Erl)
106     Resume Next
        
End Sub

Sub AgregarCompra(ByVal name As String, ByVal Desc As String)
        
        On Error GoTo AgregarCompra_Err
        

        Dim num As Integer

100     num = ComprasDonadorCheck(name)

102     Call WriteVar(CuentasPath & UCase$(name & ".act"), "COMPRAS", "CANTIDAD", num + 1)
104     Call WriteVar(CuentasPath & UCase$(name & ".act"), "COMPRAS", num + 1, Desc)

        
        Exit Sub

AgregarCompra_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModDonador.AgregarCompra", Erl)
108     Resume Next
        
End Sub

Sub RestarCreditosDonador(name As String, Cantidad As Long)
        
        On Error GoTo RestarCreditosDonador_Err
        

        Dim creditos As Long

100     Call AgregarCreditosCanjeados(name, Cantidad)

102     creditos = CreditosDonadorCheck(name) - Cantidad

104     Call WriteVar(CuentasPath & UCase$(name & ".act"), "DONADOR", "CREDITOS", creditos)

        
        Exit Sub

RestarCreditosDonador_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModDonador.RestarCreditosDonador", Erl)
108     Resume Next
        
End Sub

Sub AgregarCreditosCanjeados(name As String, Cantidad As Long)
        
        On Error GoTo AgregarCreditosCanjeados_Err
        

        Dim creditos As Long

100     creditos = CreditosCanjeadosCheck(name) + Cantidad

102     Call WriteVar(CuentasPath & UCase$(name & ".act"), "DONADOR", "CREDITOSCANJEADOS", creditos)

        
        Exit Sub

AgregarCreditosCanjeados_Err:
104     Call RegistrarError(Err.Number, Err.description, "ModDonador.AgregarCreditosCanjeados", Erl)
106     Resume Next
        
End Sub

Sub LoadDonadores()
        
        On Error GoTo LoadDonadores_Err
        

        Dim NumDonadores As Integer

        Dim tDon         As TDonador, i As Integer

100     If Not FileExist(DatPath & "Donadores.dat", vbNormal) Then Exit Sub

102     NumDonadores = val(GetVar(DatPath & "Donadores.dat", "INIT", "NumeroDonadores"))

104     For i = 1 To NumDonadores
106         Set tDon = New TDonador

108         With tDon
110             .name = GetVar(DatPath & "Donadores.dat", "DONADOR" & i, "USER")
112             .FechaExpiracion = GetVar(DatPath & "Donadores.dat", "DONADOR" & i, "FECHAEXPIRACION")
114             Call Donadores.Add(tDon)

            End With

        Next

        
        Exit Sub

LoadDonadores_Err:
116     Call RegistrarError(Err.Number, Err.description, "ModDonador.LoadDonadores", Erl)
118     Resume Next
        
End Sub

Public Function ChangeDonador(ByVal name As String, ByVal Baneado As Byte) As Boolean
        
        On Error GoTo ChangeDonador_Err
        

100     If FileExist(CuentasPath & name & ".act", vbNormal) Then
102         Call FinDonador(name)

        End If

        
        Exit Function

ChangeDonador_Err:
104     Call RegistrarError(Err.Number, Err.description, "ModDonador.ChangeDonador", Erl)
106     Resume Next
        
End Function

Public Function FinDonador(ByVal name As String) As Boolean
        
        On Error GoTo FinDonador_Err
        
       
100     Call LogearEventoDeDonador("Se finalizo suscripcion de la cuenta " & name & ".")
        'Unban the character
102     Call WriteVar(CuentasPath & name & ".act", "DONADOR", "DONADOR", "0")
104     Call WriteVar(CuentasPath & name & ".act", "DONADOR", "FECHAEXPIRACION", "")

        
        Exit Function

FinDonador_Err:
106     Call RegistrarError(Err.Number, Err.description, "ModDonador.FinDonador", Erl)
108     Resume Next
        
End Function

Public Sub LogearEventoDeDonador(Logeo As String)
        
        On Error GoTo LogearEventoDeDonador_Err
        

        Dim n As Integer

100     n = FreeFile
102     Open App.Path & "\LOGS\Donaciones.log" For Append Shared As n
104     Print #n, Date & " " & Time & " - " & Logeo
106     Close #n

        
        Exit Sub

LogearEventoDeDonador_Err:
108     Call RegistrarError(Err.Number, Err.description, "ModDonador.LogearEventoDeDonador", Erl)
110     Resume Next
        
End Sub

Public Sub CargarCodigosDonador()
        
        On Error GoTo CargarCodigosDonador_Err
        

        Dim i          As Integer

        Dim Codigostrg As String

        Dim Leer       As New clsIniReader

100     Call Leer.Initialize(App.Path & "\codigosDonadores.ini")

102     NumeroCodigos = val(Leer.GetValue("INIT", "NumCodigos"))

104     For i = 1 To NumeroCodigos
106         Codigostrg = Leer.GetValue("CODIGOS", i)
108         Codigo(i).Key = ReadField(1, Codigostrg, Asc("-"))

110         Codigo(i).Tipo = val(ReadField(2, Codigostrg, Asc("-")))

112         Codigo(i).Cantidad = val(ReadField(3, Codigostrg, Asc("-")))

114         Codigo(i).Usado = val(ReadField(4, Codigostrg, Asc("-")))

116     Next i

        
        Exit Sub

CargarCodigosDonador_Err:
118     Call RegistrarError(Err.Number, Err.description, "ModDonador.CargarCodigosDonador", Erl)
120     Resume Next
        
End Sub

Public Sub CheckearCodigo(ByVal UserIndex As Integer, ByVal CodigoKey As String)
        
        On Error GoTo CheckearCodigo_Err
        

        Dim LogCheckCodigo As String

100     LogCheckCodigo = vbCrLf & "****************************************************" & vbCrLf
102     LogCheckCodigo = LogCheckCodigo & "El usuario " & UserList(UserIndex).name & " ingresó el codigo: " & CodigoKey & "." & vbCrLf

        Dim i As Integer

104     For i = 1 To NumeroCodigos

106         If CodigoKey = Codigo(i).Key Then

                'Call WriteConsoleMsg(UserIndex, "¡Tu codigo es valido!", FontTypeNames.FONTTYPE_New_Naranja)
108             If Codigo(i).Usado = 0 Then

110                 Select Case Codigo(i).Tipo

                        Case 1 'Creditos
112                         Call AgregarCreditosDonador(UserList(UserIndex).Cuenta, CLng(Codigo(i).Cantidad))
114                         Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & Codigo(i).Cantidad & " creditos a tu cuenta. Tu saldo actual es de: " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                        
116                     Case 2 ' Tiempo
                        
118                         If DonadorCheck(UserList(UserIndex).Cuenta) = 1 Then
120                             Call DonadorTiempo(UserList(UserIndex).Cuenta, Codigo(i).Cantidad)
122                             Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & Codigo(i).Cantidad & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
124                             UserList(UserIndex).donador.activo = 1
                            Else
126                             Call DonadorTiempo(UserList(UserIndex).Cuenta, Codigo(i).Cantidad)
128                             Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya sos donador. Este benefico durara " & Codigo(i).Cantidad & " dias.", FontTypeNames.FONTTYPE_WARNING)
130                             Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & Codigo(i).Cantidad & " dias a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
132                             Call WriteConsoleMsg(UserIndex, "Te pedimos que relogees tu personaje para empezar a disfrutar los beneficios.", FontTypeNames.FONTTYPE_WARNING)

                            End If
                        
                    End Select

134                 Call WriteActShop(UserIndex)
                    '   Call WriteVar(App.Path & "\cuentas\" & name & ".act", "DONADOR", "CREDITOS", creditos)
                
136                 LogCheckCodigo = LogCheckCodigo & "El usuario " & UserList(UserIndex).name & " canjeo el codigo: " & CodigoKey & "." & vbCrLf
                
138                 Codigo(i).Usado = 1
140                 Call WriteVar(App.Path & "\codigosDonadores.ini", "CODIGOS", i, Codigo(i).Key & "-" & Codigo(i).Tipo & "-" & Codigo(i).Cantidad & "-" & Codigo(i).Usado)
142                 LogCheckCodigo = LogCheckCodigo & "****************************************************" & vbCrLf
                Else
144                 Call WriteConsoleMsg(UserIndex, "¡Ese codigo ya ha sido usado.", FontTypeNames.FONTTYPE_WARNING)
                
146                 LogCheckCodigo = LogCheckCodigo & "El codigo ya habia sido usado." & vbCrLf
148                 LogCheckCodigo = LogCheckCodigo & "****************************************************" & vbCrLf

                End If

150             Call LogearEventoDeDonador(LogCheckCodigo)
                Exit Sub
                Exit For
        
            End If

152     Next i

154     LogCheckCodigo = LogCheckCodigo & "El codigo no existe." & vbCrLf
156     LogCheckCodigo = LogCheckCodigo & "****************************************************" & vbCrLf
158     Call LogearEventoDeDonador(LogCheckCodigo)

160     Call WriteConsoleMsg(UserIndex, "¡Tu codigo es invalido!", FontTypeNames.FONTTYPE_WARNING)

        
        Exit Sub

CheckearCodigo_Err:
162     Call RegistrarError(Err.Number, Err.description, "ModDonador.CheckearCodigo", Erl)
164     Resume Next
        
End Sub
    
