Attribute VB_Name = "ModDonador"
Public Donadores As New Collection
Public Codigo(1 To 800) As CodigoData
Public NumeroCodigos As Integer

Public Type CodigoData
    Key As String
    Tipo As Byte
    Cantidad As Integer
    Usado As Byte
End Type



Option Explicit


Public Sub DonadorTiempo(ByVal nombre As String, ByVal dias As Integer)




If DonadorCheck(nombre) = 0 Then

    Dim tDon As TDonador
    
    Set tDon = New TDonador
    
    tDon.name = nombre
    tDon.FechaExpiracion = (Now + dias)
    
    
    Call Donadores.Add(tDon)
    Call SaveDonador(Donadores.Count)
    Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & nombre & " agrego " & dias & " días de donador.", FontTypeNames.FONTTYPE_New_DONADOR))
    Call LogearEventoDeDonador("Se agregaron " & dias & " a la cuenta " & nombre & ".")
Else



    'UserList(UserIndex).donador.FechaExpiracion = UserList(UserIndex).donador.FechaExpiracion + dias
    
    Dim LoopC As Integer
    For LoopC = 1 To Donadores.Count
        If UCase$(Donadores(LoopC).name) = UCase$(nombre) Then
            Donadores(LoopC).FechaExpiracion = Donadores(LoopC).FechaExpiracion + dias
            Call SaveDonador(LoopC)
            Exit For
        End If
    Next LoopC
End If


    

End Sub

Sub SaveDonadores()
Dim num As Integer
 

Call WriteVar(DatPath & "Donadores.dat", "INIT", "NumeroDonadores", val(Donadores.Count))

For num = 1 To Donadores.Count
    Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "USER", Donadores(num).name)
    Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "FECHAEXPIRACION", Donadores(num).FechaExpiracion)
Next

End Sub
Sub SaveDonador(num As Integer)

Call WriteVar(DatPath & "Donadores.dat", "INIT", "NumeroDonadores", Donadores.Count)
Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "USER", Donadores(num).name)
Call WriteVar(DatPath & "Donadores.dat", "DONADOR" & num, "FECHAEXPIRACION", Donadores(num).FechaExpiracion)


Call WriteVar(CuentasPath & Donadores(num).name & ".act", "DONADOR", "DONADOR", "1")
Call WriteVar(CuentasPath & Donadores(num).name & ".act", "DONADOR", "FECHAEXPIRACION", Donadores(num).FechaExpiracion)
End Sub
Sub AgregarCreditosDonador(name As String, Cantidad As Long)

Dim creditos As Long
creditos = CreditosDonadorCheck(name) + Cantidad

'Call LogearEventoDeDonador("Se agregaron " & Cantidad & " creditos a la cuenta " & Name & ".")
Call WriteVar(CuentasPath & UCase$(name & ".act"), "DONADOR", "CREDITOS", creditos)
'Call AgregarCompra(Name, Date & " - Se agregaron " & Cantidad & " creditos a la cuenta " & Name & ".")
End Sub
Sub AgregarCompra(ByVal name As String, ByVal Desc As String)
Dim num As Integer
num = ComprasDonadorCheck(name)

Call WriteVar(CuentasPath & UCase$(name & ".act"), "COMPRAS", "CANTIDAD", num + 1)
Call WriteVar(CuentasPath & UCase$(name & ".act"), "COMPRAS", num + 1, Desc)

End Sub
Sub RestarCreditosDonador(name As String, Cantidad As Long)

Dim creditos As Long

Call AgregarCreditosCanjeados(name, Cantidad)

creditos = CreditosDonadorCheck(name) - Cantidad


Call WriteVar(CuentasPath & UCase$(name & ".act"), "DONADOR", "CREDITOS", creditos)

End Sub
Sub AgregarCreditosCanjeados(name As String, Cantidad As Long)

Dim creditos As Long
creditos = CreditosCanjeadosCheck(name) + Cantidad


Call WriteVar(CuentasPath & UCase$(name & ".act"), "DONADOR", "CREDITOSCANJEADOS", creditos)

End Sub
Sub LoadDonadores()
Dim NumDonadores As Integer
Dim tDon As TDonador, i As Integer

If Not FileExist(DatPath & "Donadores.dat", vbNormal) Then Exit Sub

NumDonadores = val(GetVar(DatPath & "Donadores.dat", "INIT", "NumeroDonadores"))

For i = 1 To NumDonadores
    Set tDon = New TDonador
    With tDon
        .name = GetVar(DatPath & "Donadores.dat", "DONADOR" & i, "USER")
        .FechaExpiracion = GetVar(DatPath & "Donadores.dat", "DONADOR" & i, "FECHAEXPIRACION")
        Call Donadores.Add(tDon)
    End With
Next

End Sub


Public Function ChangeDonador(ByVal name As String, ByVal Baneado As Byte) As Boolean

If FileExist(CuentasPath & name & ".act", vbNormal) Then
        Call FinDonador(name)
End If


    

End Function


Public Function FinDonador(ByVal name As String) As Boolean
       
Call LogearEventoDeDonador("Se finalizo suscripcion de la cuenta " & name & ".")
'Unban the character
Call WriteVar(CuentasPath & name & ".act", "DONADOR", "DONADOR", "0")
Call WriteVar(CuentasPath & name & ".act", "DONADOR", "FECHAEXPIRACION", "")

End Function

Public Sub LogearEventoDeDonador(Logeo As String)
Dim n As Integer
        n = FreeFile
        Open App.Path & "\LOGS\Donaciones.log" For Append Shared As n
        Print #n, Date & " " & Time & " - " & Logeo
        Close #n
End Sub

Public Sub CargarCodigosDonador()
Dim i As Integer

Dim Codigostrg As String


Dim Leer As New clsIniReader

Call Leer.Initialize(App.Path & "\codigosDonadores.ini")

NumeroCodigos = val(Leer.GetValue("INIT", "NumCodigos"))

For i = 1 To NumeroCodigos
    Codigostrg = Leer.GetValue("CODIGOS", i)
    Codigo(i).Key = ReadField(1, Codigostrg, Asc("-"))

    Codigo(i).Tipo = val(ReadField(2, Codigostrg, Asc("-")))

    Codigo(i).Cantidad = val(ReadField(3, Codigostrg, Asc("-")))

    Codigo(i).Usado = val(ReadField(4, Codigostrg, Asc("-")))

Next i

End Sub

Public Sub CheckearCodigo(ByVal UserIndex As Integer, ByVal CodigoKey As String)

Dim LogCheckCodigo As String

LogCheckCodigo = vbCrLf & "****************************************************" & vbCrLf
LogCheckCodigo = LogCheckCodigo & "El usuario " & UserList(UserIndex).name & " ingresó el codigo: " & CodigoKey & "." & vbCrLf

Dim i As Integer

For i = 1 To NumeroCodigos
    If CodigoKey = Codigo(i).Key Then
        'Call WriteConsoleMsg(UserIndex, "¡Tu codigo es valido!", FontTypeNames.FONTTYPE_New_Naranja)
            If Codigo(i).Usado = 0 Then
                Select Case Codigo(i).Tipo
                    Case 1 'Creditos
                        Call AgregarCreditosDonador(UserList(UserIndex).Cuenta, CLng(Codigo(i).Cantidad))
                        Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & Codigo(i).Cantidad & " creditos a tu cuenta. Tu saldo actual es de: " & CreditosDonadorCheck(UserList(UserIndex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                        
                    Case 2 ' Tiempo
                        
                        If DonadorCheck(UserList(UserIndex).Cuenta) = 1 Then
                            Call DonadorTiempo(UserList(UserIndex).Cuenta, Codigo(i).Cantidad)
                            Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & Codigo(i).Cantidad & " dias de donador a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
                            UserList(UserIndex).donador.activo = 1
                        Else
                            Call DonadorTiempo(UserList(UserIndex).Cuenta, Codigo(i).Cantidad)
                            Call WriteConsoleMsg(UserIndex, "¡Felicitaciones! Ya sos donador. Este benefico durara " & Codigo(i).Cantidad & " dias.", FontTypeNames.FONTTYPE_WARNING)
                            Call WriteConsoleMsg(UserIndex, "¡Se han añadido " & Codigo(i).Cantidad & " dias a tu cuenta.", FontTypeNames.FONTTYPE_WARNING)
                            Call WriteConsoleMsg(UserIndex, "Te pedimos que relogees tu personaje para empezar a disfrutar los beneficios.", FontTypeNames.FONTTYPE_WARNING)
                        End If
                        
                End Select
                Call WriteActShop(UserIndex)
             '   Call WriteVar(App.Path & "\cuentas\" & name & ".act", "DONADOR", "CREDITOS", creditos)
                
                LogCheckCodigo = LogCheckCodigo & "El usuario " & UserList(UserIndex).name & " canjeo el codigo: " & CodigoKey & "." & vbCrLf
                
                Codigo(i).Usado = 1
                Call WriteVar(App.Path & "\codigosDonadores.ini", "CODIGOS", i, Codigo(i).Key & "-" & Codigo(i).Tipo & "-" & Codigo(i).Cantidad & "-" & Codigo(i).Usado)
                LogCheckCodigo = LogCheckCodigo & "****************************************************" & vbCrLf
            Else
                Call WriteConsoleMsg(UserIndex, "¡Ese codigo ya ha sido usado.", FontTypeNames.FONTTYPE_WARNING)
                
                LogCheckCodigo = LogCheckCodigo & "El codigo ya habia sido usado." & vbCrLf
                LogCheckCodigo = LogCheckCodigo & "****************************************************" & vbCrLf
            End If
            Call LogearEventoDeDonador(LogCheckCodigo)
        Exit Sub
        Exit For
        
    End If
Next i


LogCheckCodigo = LogCheckCodigo & "El codigo no existe." & vbCrLf
LogCheckCodigo = LogCheckCodigo & "****************************************************" & vbCrLf
Call LogearEventoDeDonador(LogCheckCodigo)


Call WriteConsoleMsg(UserIndex, "¡Tu codigo es invalido!", FontTypeNames.FONTTYPE_WARNING)
        



End Sub


    
