Attribute VB_Name = "SecurityIp"
'**************************************************************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y dise�ado por DuNga (ltourrilhes@gmail.com)
'**************************************************************
Option Explicit

'*************************************************  *************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y dise�ado por DuNga (ltourrilhes@gmail.com)
'*************************************************  *************

Private IpTables()                     As Long 'USAMOS 2 LONGS: UNO DE LA IP, SEGUIDO DE UNO DE LA INFO

Private EntrysCounter                  As Long

Private MaxValue                       As Long

Private Multiplicado                   As Long 'Cuantas veces multiplike el EntrysCounter para que me entren?

Private Const IntervaloEntreConexiones As Long = 500

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declaraciones para maximas conexiones por usuario
'Agregado por EL OSO
Private MaxConTables()                 As Long

Private MaxConTablesEntry              As Long     'puntero a la ultima insertada

Private Const LIMITECONEXIONESxIP      As Long = 10

Private Enum e_SecurityIpTabla

    IP_INTERVALOS = 1
    IP_LIMITECONEXIONES = 2

End Enum

Public Sub InitIpTables(ByVal OptCountersValue As Long)
    '*************************************************  *************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: EL OSO 21/01/06. Soporte para MaxConTables
    '
    '*************************************************  *************
    EntrysCounter = OptCountersValue
    Multiplicado = 1

    ReDim IpTables(EntrysCounter * 2) As Long
    MaxValue = 0

    ReDim MaxConTables(Declaraciones.MaxUsers * 2 - 1) As Long
    MaxConTablesEntry = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''FUNCIONES PARA INTERVALOS'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub IpSecurityMantenimientoLista()
    '*************************************************  *************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '*************************************************  *************
    'Las borro todas cada 1 hora, asi se "renuevan"
    EntrysCounter = EntrysCounter \ Multiplicado
    Multiplicado = 1
    ReDim IpTables(EntrysCounter * 2) As Long
    MaxValue = 0

End Sub

Public Function IpSecurityAceptarNuevaConexion(ByVal ip As Long) As Boolean

    '*************************************************  *************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '*************************************************  *************
    Dim IpTableIndex As Long

    IpTableIndex = FindTableIp(ip, IP_INTERVALOS)
    
    If IpTableIndex >= 0 Then
        If IpTables(IpTableIndex + 1) + IntervaloEntreConexiones <= GetTickCount Then   'No est� saturando de connects?
            IpTables(IpTableIndex + 1) = GetTickCount
            IpSecurityAceptarNuevaConexion = True
            Debug.Print "CONEXION ACEPTADA"
            Exit Function
        Else
            IpSecurityAceptarNuevaConexion = False

            Debug.Print "CONEXION NO ACEPTADA"
            Exit Function

        End If

    Else
        IpTableIndex = Not IpTableIndex
        AddNewIpIntervalo ip, IpTableIndex
        IpTables(IpTableIndex + 1) = GetTickCount
        IpSecurityAceptarNuevaConexion = True
        Exit Function

    End If

End Function

Private Sub AddNewIpIntervalo(ByVal ip As Long, ByVal Index As Long)

    '*************************************************  *************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '*************************************************  *************
    '2) Pruebo si hay espacio, sino agrando la lista
    If MaxValue + 1 > EntrysCounter Then
        EntrysCounter = EntrysCounter \ Multiplicado
        Multiplicado = Multiplicado + 1
        EntrysCounter = EntrysCounter * Multiplicado
        
        ReDim Preserve IpTables(EntrysCounter * 2) As Long

    End If
    
    '4) Corro todo el array para arriba
    Call CopyMemory(IpTables(Index + 2), IpTables(Index), (MaxValue - Index \ 2) * 8)   '*4 (peso del long) * 2(cantidad de elementos por c/u)
    IpTables(Index) = ip
    
    '3) Subo el indicador de el maximo valor almacenado y listo :)
    MaxValue = MaxValue + 1

End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''FUNCIONES PARA LIMITES X IP''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IPSecuritySuperaLimiteConexiones(ByVal ip As Long) As Boolean

    Dim IpTableIndex As Long

    IpTableIndex = FindTableIp(ip, IP_LIMITECONEXIONES)
    
    If IpTableIndex >= 0 Then
        
        If MaxConTables(IpTableIndex + 1) < LIMITECONEXIONESxIP Then
            LogIP ("Agregamos conexion a " & ip & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            Debug.Print "suma conexion a " & ip & " total " & MaxConTables(IpTableIndex + 1) + 1
            MaxConTables(IpTableIndex + 1) = MaxConTables(IpTableIndex + 1) + 1
            IPSecuritySuperaLimiteConexiones = False
        Else
            LogIP ("rechazamos conexion de " & ip & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
            Debug.Print "rechaza conexion a " & ip
            IPSecuritySuperaLimiteConexiones = True

        End If

    Else
        IPSecuritySuperaLimiteConexiones = False

        If MaxConTablesEntry < Declaraciones.MaxUsers Then  'si hay espacio..
            IpTableIndex = Not IpTableIndex
            AddNewIpLimiteConexiones ip, IpTableIndex    'iptableindex es donde lo agrego
            MaxConTables(IpTableIndex + 1) = 1
        Else
            Call LogCriticEvent("SecurityIP.IPSecuritySuperaLimiteConexiones: Se supero la disponibilidad de slots.")

        End If

    End If

End Function

Private Sub AddNewIpLimiteConexiones(ByVal ip As Long, ByVal Index As Long)
    '*************************************************  *************
    'Author: (EL OSO)
    'Last Modify Date: Unknow
    '
    '*************************************************  *************
    'Debug.Print "agrega conexion a " & ip
    'Debug.Print "(Declaraciones.MaxUsers - index) = " & (Declaraciones.MaxUsers - Index)
    '4) Corro todo el array para arriba
    'Call CopyMemory(MaxConTables(Index + 2), MaxConTables(Index), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
    'MaxConTables(Index) = ip

    '3) Subo el indicador de el maximo valor almacenado y listo :)
    'MaxConTablesEntry = MaxConTablesEntry + 1

    '*************************************************    *************
    'Author: (EL OSO)
    'Last Modify Date: 16/2/2006
    'Modified by Juan Mart�n Sotuyo Dodero (Maraxus)
    '*************************************************    *************
    Debug.Print "agrega conexion a " & ip
    Debug.Print "(Declaraciones.MaxUsers - index) = " & (Declaraciones.MaxUsers - Index)
    Debug.Print "Agrega conexion a nueva IP " & ip

    '4) Corro todo el array para arriba
    Dim temp() As Long

    ReDim temp((MaxConTablesEntry - Index \ 2) * 2) As Long  'VB no deja inicializar con rangos variables...
    Call CopyMemory(temp(0), MaxConTables(Index), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
    Call CopyMemory(MaxConTables(Index + 2), temp(0), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
    MaxConTables(Index) = ip

    '3) Subo el indicador de el maximo valor almacenado y listo :)
    MaxConTablesEntry = MaxConTablesEntry + 1

End Sub

Public Sub IpRestarConexion(ByVal ip As Long)

    Dim Key As Long

    Debug.Print "resta conexion a " & ip
    
    Key = FindTableIp(ip, IP_LIMITECONEXIONES)
    
    If Key >= 0 Then
        If MaxConTables(Key + 1) > 0 Then
            MaxConTables(Key + 1) = MaxConTables(Key + 1) - 1

        End If

        Call LogIP("restamos conexion a " & ip & " key=" & Key & ". Conexiones: " & MaxConTables(Key + 1))

        If MaxConTables(Key + 1) <= 0 Then
            'la limpiamos
            Call CopyMemory(MaxConTables(Key), MaxConTables(Key + 2), (MaxConTablesEntry - (Key \ 2) + 1) * 8)
            MaxConTablesEntry = MaxConTablesEntry - 1

        End If

    Else 'Key <= 0
        Call LogIP("restamos conexion a " & ip & " key=" & Key & ". NEGATIVO!!")

        'LogCriticEvent "SecurityIp.IpRestarconexion obtuvo un valor negativo en key"
    End If

End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''''''FUNCIONES GENERALES''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function FindTableIp(ByVal ip As Long, ByVal Tabla As e_SecurityIpTabla) As Long

    '*************************************************  *************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    'Modified by Juan Mart�n Sotuyo Dodero (Maraxus) to use Binary Insertion
    '*************************************************  *************
    Dim First  As Long

    Dim Last   As Long

    Dim Middle As Long
    
    Select Case Tabla

        Case e_SecurityIpTabla.IP_INTERVALOS
            First = 0
            Last = MaxValue

            Do While First <= Last
                Middle = (First + Last) \ 2
                
                If (IpTables(Middle * 2) < ip) Then
                    First = Middle + 1
                ElseIf (IpTables(Middle * 2) > ip) Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function

                End If

            Loop
            FindTableIp = Not (Middle * 2)
        
        Case e_SecurityIpTabla.IP_LIMITECONEXIONES
            
            First = 0
            Last = MaxConTablesEntry

            Do While First <= Last
                Middle = (First + Last) \ 2

                If MaxConTables(Middle * 2) < ip Then
                    First = Middle + 1
                ElseIf MaxConTables(Middle * 2) > ip Then
                    Last = Middle - 1
                Else
                    FindTableIp = Middle * 2
                    Exit Function

                End If

            Loop
            FindTableIp = Not (Middle * 2)

    End Select

End Function

Public Function DumpTables()

    Dim i As Integer

    For i = 0 To MaxConTablesEntry * 2 - 1 Step 2
        Call LogCriticEvent(GetAscIP(MaxConTables(i)) & " > " & MaxConTables(i + 1))
    Next i

End Function
