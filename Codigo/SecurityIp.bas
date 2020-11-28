Attribute VB_Name = "SecurityIp"
'**************************************************************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y diseñado por DuNga (ltourrilhes@gmail.com)
'**************************************************************
Option Explicit

'*************************************************  *************
' General_IpSecurity.Bas - Maneja la seguridad de las IPs
'
' Escrito y diseñado por DuNga (ltourrilhes@gmail.com)
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
        
        On Error GoTo InitIpTables_Err
        
100     EntrysCounter = OptCountersValue
102     Multiplicado = 1

104     ReDim IpTables(EntrysCounter * 2) As Long
106     MaxValue = 0

108     ReDim MaxConTables(Declaraciones.MaxUsers * 2 - 1) As Long
110     MaxConTablesEntry = 0

        
        Exit Sub

InitIpTables_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.InitIpTables", Erl)
        Resume Next
        
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
        
        On Error GoTo IpSecurityMantenimientoLista_Err
        
100     EntrysCounter = EntrysCounter \ Multiplicado
102     Multiplicado = 1
104     ReDim IpTables(EntrysCounter * 2) As Long
106     MaxValue = 0

        
        Exit Sub

IpSecurityMantenimientoLista_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.IpSecurityMantenimientoLista", Erl)
        Resume Next
        
End Sub

Public Function IpSecurityAceptarNuevaConexion(ByVal ip As Long) As Boolean
        
        On Error GoTo IpSecurityAceptarNuevaConexion_Err
        

        '*************************************************  *************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '*************************************************  *************
        Dim IpTableIndex As Long

100     IpTableIndex = FindTableIp(ip, IP_INTERVALOS)
    
102     If IpTableIndex >= 0 Then
104         If IpTables(IpTableIndex + 1) + IntervaloEntreConexiones <= GetTickCount Then   'No está saturando de connects?
106             IpTables(IpTableIndex + 1) = GetTickCount
108             IpSecurityAceptarNuevaConexion = True
110             Debug.Print "CONEXION ACEPTADA"
                Exit Function
            Else
112             IpSecurityAceptarNuevaConexion = False

114             Debug.Print "CONEXION NO ACEPTADA"
                Exit Function

            End If

        Else
116         IpTableIndex = Not IpTableIndex
118         AddNewIpIntervalo ip, IpTableIndex
120         IpTables(IpTableIndex + 1) = GetTickCount
122         IpSecurityAceptarNuevaConexion = True
            Exit Function

        End If

        
        Exit Function

IpSecurityAceptarNuevaConexion_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.IpSecurityAceptarNuevaConexion", Erl)
        Resume Next
        
End Function

Private Sub AddNewIpIntervalo(ByVal ip As Long, ByVal Index As Long)
        
        On Error GoTo AddNewIpIntervalo_Err
        

        '*************************************************  *************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        '
        '*************************************************  *************
        '2) Pruebo si hay espacio, sino agrando la lista
100     If MaxValue + 1 > EntrysCounter Then
102         EntrysCounter = EntrysCounter \ Multiplicado
104         Multiplicado = Multiplicado + 1
106         EntrysCounter = EntrysCounter * Multiplicado
        
108         ReDim Preserve IpTables(EntrysCounter * 2) As Long

        End If
    
        '4) Corro todo el array para arriba
110     Call CopyMemory(IpTables(Index + 2), IpTables(Index), (MaxValue - Index \ 2) * 8)   '*4 (peso del long) * 2(cantidad de elementos por c/u)
112     IpTables(Index) = ip
    
        '3) Subo el indicador de el maximo valor almacenado y listo :)
114     MaxValue = MaxValue + 1

        
        Exit Sub

AddNewIpIntervalo_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.AddNewIpIntervalo", Erl)
        Resume Next
        
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''FUNCIONES PARA LIMITES X IP''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IPSecuritySuperaLimiteConexiones(ByVal ip As Long) As Boolean
        
        On Error GoTo IPSecuritySuperaLimiteConexiones_Err
        

        Dim IpTableIndex As Long

100     IpTableIndex = FindTableIp(ip, IP_LIMITECONEXIONES)
    
102     If IpTableIndex >= 0 Then
        
104         If MaxConTables(IpTableIndex + 1) < LIMITECONEXIONESxIP Then
106             LogIP ("Agregamos conexion a " & ip & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
108             Debug.Print "suma conexion a " & ip & " total " & MaxConTables(IpTableIndex + 1) + 1
110             MaxConTables(IpTableIndex + 1) = MaxConTables(IpTableIndex + 1) + 1
112             IPSecuritySuperaLimiteConexiones = False
            Else
114             LogIP ("rechazamos conexion de " & ip & " iptableindex=" & IpTableIndex & ". Conexiones: " & MaxConTables(IpTableIndex + 1))
116             Debug.Print "rechaza conexion a " & ip
118             IPSecuritySuperaLimiteConexiones = True

            End If

        Else
120         IPSecuritySuperaLimiteConexiones = False

122         If MaxConTablesEntry < Declaraciones.MaxUsers Then  'si hay espacio..
124             IpTableIndex = Not IpTableIndex
126             AddNewIpLimiteConexiones ip, IpTableIndex    'iptableindex es donde lo agrego
128             MaxConTables(IpTableIndex + 1) = 1
            Else
130             Call LogCriticEvent("SecurityIP.IPSecuritySuperaLimiteConexiones: Se supero la disponibilidad de slots.")

            End If

        End If

        
        Exit Function

IPSecuritySuperaLimiteConexiones_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.IPSecuritySuperaLimiteConexiones", Erl)
        Resume Next
        
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
        
        On Error GoTo AddNewIpLimiteConexiones_Err
        

        '3) Subo el indicador de el maximo valor almacenado y listo :)
        'MaxConTablesEntry = MaxConTablesEntry + 1

        '*************************************************    *************
        'Author: (EL OSO)
        'Last Modify Date: 16/2/2006
        'Modified by Juan Martín Sotuyo Dodero (Maraxus)
        '*************************************************    *************
100     Debug.Print "agrega conexion a " & ip
102     Debug.Print "(Declaraciones.MaxUsers - index) = " & (Declaraciones.MaxUsers - Index)
104     Debug.Print "Agrega conexion a nueva IP " & ip

        '4) Corro todo el array para arriba
        Dim temp() As Long

106     ReDim temp((MaxConTablesEntry - Index \ 2) * 2) As Long  'VB no deja inicializar con rangos variables...
108     Call CopyMemory(temp(0), MaxConTables(Index), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
110     Call CopyMemory(MaxConTables(Index + 2), temp(0), (MaxConTablesEntry - Index \ 2) * 8)    '*4 (peso del long) * 2(cantidad de elementos por c/u)
112     MaxConTables(Index) = ip

        '3) Subo el indicador de el maximo valor almacenado y listo :)
114     MaxConTablesEntry = MaxConTablesEntry + 1

        
        Exit Sub

AddNewIpLimiteConexiones_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.AddNewIpLimiteConexiones", Erl)
        Resume Next
        
End Sub

Public Sub IpRestarConexion(ByVal ip As Long)
        
        On Error GoTo IpRestarConexion_Err
        

        Dim Key As Long

100     Debug.Print "resta conexion a " & ip
    
102     Key = FindTableIp(ip, IP_LIMITECONEXIONES)
    
104     If Key >= 0 Then
106         If MaxConTables(Key + 1) > 0 Then
108             MaxConTables(Key + 1) = MaxConTables(Key + 1) - 1

            End If

110         Call LogIP("restamos conexion a " & ip & " key=" & Key & ". Conexiones: " & MaxConTables(Key + 1))

112         If MaxConTables(Key + 1) <= 0 Then
                'la limpiamos
114             Call CopyMemory(MaxConTables(Key), MaxConTables(Key + 2), (MaxConTablesEntry - (Key \ 2) + 1) * 8)
116             MaxConTablesEntry = MaxConTablesEntry - 1

            End If

        Else 'Key <= 0
118         Call LogIP("restamos conexion a " & ip & " key=" & Key & ". NEGATIVO!!")

            'LogCriticEvent "SecurityIp.IpRestarconexion obtuvo un valor negativo en key"
        End If

        
        Exit Sub

IpRestarConexion_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.IpRestarConexion", Erl)
        Resume Next
        
End Sub

' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ''''''''''''''''''''''''FUNCIONES GENERALES''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function FindTableIp(ByVal ip As Long, ByVal Tabla As e_SecurityIpTabla) As Long
        
        On Error GoTo FindTableIp_Err
        

        '*************************************************  *************
        'Author: Lucio N. Tourrilhes (DuNga)
        'Last Modify Date: Unknow
        'Modified by Juan Martín Sotuyo Dodero (Maraxus) to use Binary Insertion
        '*************************************************  *************
        Dim First  As Long

        Dim Last   As Long

        Dim Middle As Long
    
100     Select Case Tabla

            Case e_SecurityIpTabla.IP_INTERVALOS
102             First = 0
104             Last = MaxValue

106             Do While First <= Last
108                 Middle = (First + Last) \ 2
                
110                 If (IpTables(Middle * 2) < ip) Then
112                     First = Middle + 1
114                 ElseIf (IpTables(Middle * 2) > ip) Then
116                     Last = Middle - 1
                    Else
118                     FindTableIp = Middle * 2
                        Exit Function

                    End If

                Loop
120             FindTableIp = Not (Middle * 2)
        
122         Case e_SecurityIpTabla.IP_LIMITECONEXIONES
            
124             First = 0
126             Last = MaxConTablesEntry

128             Do While First <= Last
130                 Middle = (First + Last) \ 2

132                 If MaxConTables(Middle * 2) < ip Then
134                     First = Middle + 1
136                 ElseIf MaxConTables(Middle * 2) > ip Then
138                     Last = Middle - 1
                    Else
140                     FindTableIp = Middle * 2
                        Exit Function

                    End If

                Loop
142             FindTableIp = Not (Middle * 2)

        End Select

        
        Exit Function

FindTableIp_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.FindTableIp", Erl)
        Resume Next
        
End Function

Public Function DumpTables()
        
        On Error GoTo DumpTables_Err
        

        Dim i As Integer

100     For i = 0 To MaxConTablesEntry * 2 - 1 Step 2
102         Call LogCriticEvent(GetAscIP(MaxConTables(i)) & " > " & MaxConTables(i + 1))
104     Next i

        
        Exit Function

DumpTables_Err:
        Call RegistrarError(Err.Number, Err.description, "SecurityIp.DumpTables", Erl)
        Resume Next
        
End Function
