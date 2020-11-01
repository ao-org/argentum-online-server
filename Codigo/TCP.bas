Attribute VB_Name = "TCP"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpo(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewBody As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(UserIndex).genero
UserRaza = UserList(UserIndex).raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Enano
                NewBody = 300
            Case eRaza.Gnomo
                NewBody = 300
            Case eRaza.Orco
                NewBody = 582
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Gnomo
                NewBody = 300
            Case eRaza.Enano
                NewBody = 300
            Case eRaza.Orco
                NewBody = 581
        End Select
End Select
UserList(UserIndex).Char.Body = NewBody
End Sub
Sub AsignarAtributos(ByVal UserIndex As String)
    Select Case UserList(UserIndex).raza
        Case eRaza.Humano
            UserList(UserIndex).Stats.UserAtributos(1) = 19
            UserList(UserIndex).Stats.UserAtributos(2) = 19
            UserList(UserIndex).Stats.UserAtributos(3) = 19
            UserList(UserIndex).Stats.UserAtributos(4) = 20
         Case eRaza.Elfo
            UserList(UserIndex).Stats.UserAtributos(1) = 18
            UserList(UserIndex).Stats.UserAtributos(2) = 20
            UserList(UserIndex).Stats.UserAtributos(3) = 21
            UserList(UserIndex).Stats.UserAtributos(4) = 18
        Case eRaza.Drow
            UserList(UserIndex).Stats.UserAtributos(1) = 20
            UserList(UserIndex).Stats.UserAtributos(2) = 18
            UserList(UserIndex).Stats.UserAtributos(3) = 20
            UserList(UserIndex).Stats.UserAtributos(4) = 19
        Case eRaza.Gnomo
            UserList(UserIndex).Stats.UserAtributos(1) = 13
            UserList(UserIndex).Stats.UserAtributos(2) = 21
            UserList(UserIndex).Stats.UserAtributos(3) = 22
            UserList(UserIndex).Stats.UserAtributos(4) = 17
        Case eRaza.Enano
            UserList(UserIndex).Stats.UserAtributos(1) = 21
            UserList(UserIndex).Stats.UserAtributos(2) = 17
            UserList(UserIndex).Stats.UserAtributos(3) = 12
            UserList(UserIndex).Stats.UserAtributos(4) = 22
        Case eRaza.Orco
            UserList(UserIndex).Stats.UserAtributos(1) = 23
            UserList(UserIndex).Stats.UserAtributos(2) = 17
            UserList(UserIndex).Stats.UserAtributos(3) = 12
            UserList(UserIndex).Stats.UserAtributos(4) = 21
    End Select
End Sub
Sub RellenarInventario(ByVal UserIndex As String)

    With UserList(UserIndex)
        
        Dim NumItems As Integer
        NumItems = 1
    
        ' Todos reciben pociones rojas
        .Invent.Object(NumItems).ObjIndex = 1616 'Pocion Roja
        .Invent.Object(NumItems).Amount = 100
        NumItems = NumItems + 1
        
        ' Magicas puras reciben más azules
        Select Case .clase
            Case eClass.Mage, eClass.Druid
                .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
                .Invent.Object(NumItems).Amount = 100
                NumItems = NumItems + 1
        End Select
        
        ' Semi mágicas reciben menos
        Select Case .clase
            Case eClass.Bard, eClass.Cleric, eClass.Paladin, eClass.Assasin
                .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
                .Invent.Object(NumItems).Amount = 50
                NumItems = NumItems + 1
        End Select

        ' Arma y hechizos
        Select Case .clase
            Case eClass.Mage
                .Invent.Object(NumItems).ObjIndex = 1356 ' Báculo (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1

                .Stats.UserHechizos(1) = 1 ' Proyectil
                .Stats.UserHechizos(2) = 2 ' Saeta
                .Stats.UserHechizos(3) = 11 ' Curar Veneno

            Case eClass.Bard
                .Invent.Object(NumItems).ObjIndex = 1623 ' Nudillos oxidados (Newbies)
                .Invent.NudilloSlot = NumItems
                NumItems = NumItems + 1

                .Stats.UserHechizos(1) = 1 ' Proyectil
                .Stats.UserHechizos(2) = 2 ' Saeta
                .Stats.UserHechizos(3) = 11 ' Curar Veneno

            Case eClass.Druid
                .Invent.Object(NumItems).ObjIndex = 420 ' Daga (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1

                .Stats.UserHechizos(1) = 1 ' Proyectil
                .Stats.UserHechizos(2) = 2 ' Saeta
                .Stats.UserHechizos(3) = 11 ' Curar Veneno
                .Stats.UserHechizos(4) = 12 ' Heridas Leves

            Case eClass.Assasin
                .Invent.Object(NumItems).ObjIndex = 420 ' Daga (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1

                .Stats.UserHechizos(1) = 1 ' Proyectil
                .Stats.UserHechizos(2) = 2 ' Saeta
                .Stats.UserHechizos(3) = 11 ' Curar Veneno

            Case eClass.Cleric
                .Invent.Object(NumItems).ObjIndex = 2085 ' Espada Larga (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1

                .Stats.UserHechizos(1) = 1 ' Proyectil
                .Stats.UserHechizos(2) = 2 ' Saeta
                .Stats.UserHechizos(3) = 11 ' Curar Veneno
                .Stats.UserHechizos(4) = 12 ' Heridas Leves

            Case eClass.Paladin
                .Invent.Object(NumItems).ObjIndex = 2085 ' Espada Larga (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1
                
                .Stats.UserHechizos(1) = 1 'Proyectil
                .Stats.UserHechizos(2) = 2 'Saeta
                .Stats.UserHechizos(3) = 11 'Curar Veneno

            Case eClass.Hunter
                .Invent.Object(NumItems).ObjIndex = 1355 ' Arco Simple (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1
                
                .Invent.Object(NumItems).ObjIndex = 1357 ' Flecha (Newbies)
                .Invent.Object(NumItems).Amount = 2000
                .Invent.Object(NumItems).Equipped = 1
                .Invent.MunicionEqpSlot = NumItems
                .Invent.MunicionEqpObjIndex = .Invent.Object(NumItems).ObjIndex
                NumItems = NumItems + 1

            Case eClass.Warrior
                .Invent.Object(NumItems).ObjIndex = 2085 ' Espada Larga (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1

            Case eClass.Trabajador
               .Invent.Object(NumItems).ObjIndex = 2085 ' Espada Larga (Newbies)
                .Invent.WeaponEqpSlot = NumItems
                NumItems = NumItems + 1
        End Select
        
        ' Pociones amarillas y verdes
        Select Case .clase
            Case eClass.Assasin, eClass.Bard, eClass.Cleric, eClass.Hunter, eClass.Paladin, eClass.Trabajador, eClass.Warrior
                .Invent.Object(NumItems).ObjIndex = 1618 ' Pocion Amarilla
                .Invent.Object(NumItems).Amount = 25
                NumItems = NumItems + 1

                .Invent.Object(NumItems).ObjIndex = 1619 ' Pocion Verde
                .Invent.Object(NumItems).Amount = 25
                NumItems = NumItems + 1
        End Select
        
        ' Equipo el arma
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.Object(.Invent.WeaponEqpSlot).Amount = 1
            .Invent.Object(.Invent.WeaponEqpSlot).Equipped = 1
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
            .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
        ' O los nudillos
        ElseIf .Invent.NudilloSlot > 0 Then
            .Invent.Object(.Invent.NudilloSlot).Amount = 1
            .Invent.Object(.Invent.NudilloSlot).Equipped = 1
            .Invent.NudilloObjIndex = .Invent.Object(.Invent.NudilloSlot).ObjIndex
            .Char.WeaponAnim = ObjData(.Invent.NudilloObjIndex).WeaponAnim
        End If
        
        ' Vestimenta común
        .Invent.Object(NumItems).ObjIndex = 1622 ' Vestimenta Comun
        .Invent.Object(NumItems).Amount = 1
        .Invent.Object(NumItems).Equipped = 1
        .Invent.ArmourEqpSlot = NumItems
        .Invent.ArmourEqpObjIndex = .Invent.Object(NumItems).ObjIndex
        NumItems = NumItems + 1

        ' Animación según raza
        If .raza = Enano Or .raza = Gnomo Then
            .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).RopajeBajo
        Else
            .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        End If
        
        ' Comida y bebida
        .Invent.Object(NumItems).ObjIndex = 573 ' Manzana
        .Invent.Object(NumItems).Amount = 100
        NumItems = NumItems + 1

        .Invent.Object(NumItems).ObjIndex = 572 ' Agua
        .Invent.Object(NumItems).Amount = 100
        NumItems = NumItems + 1

        ' Seteo la cantidad de items
        .Invent.NroItems = NumItems

    End With

   
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByVal Head As Integer, ByRef UserCuenta As String)
    '*************************************************
    'Author: Unknown
    'Last modified: 20/4/2007
    'Conecta un nuevo Usuario
    '23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
    '24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
    '12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
    '20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
    '09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
    '*************************************************
    
    If Not AsciiValidos(name) Or LenB(name) = 0 Then
        Call WriteErrorMsg(UserIndex, "Nombre invalido.")
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado crear a " & name & " desde la IP " & UserList(UserIndex).ip)
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        Exit Sub
    End If
    
    Dim LoopC As Long
    
    '¿Existe el personaje?
    If PersonajeExiste(name) Then
        Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
        Exit Sub
    End If
    
    'Prevenimos algun bug con dados inválidos
    If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then Exit Sub
    
    UserList(UserIndex).flags.Muerto = 0
    UserList(UserIndex).flags.Escondido = 0

    UserList(UserIndex).flags.Casado = 0
    UserList(UserIndex).flags.Pareja = ""

    UserList(UserIndex).name = name
    UserList(UserIndex).clase = UserClase
    UserList(UserIndex).raza = UserRaza
    
    UserList(UserIndex).Char.Head = Head
    
    UserList(UserIndex).genero = UserSexo
    UserList(UserIndex).Hogar = 1
    
    
    '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
    UserList(UserIndex).Stats.SkillPts = 10
    
    
    UserList(UserIndex).Char.heading = eHeading.SOUTH
    
    Call DarCuerpo(UserIndex) 'Ladder REVISAR
    
    UserList(UserIndex).OrigChar = UserList(UserIndex).Char

    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco

    'Call AsignarAtributos(UserIndex)

    UserList(UserIndex).Stats.MaxHp = ModVida(UserList(UserIndex).raza).Inicial(UserList(UserIndex).clase)
    UserList(UserIndex).Stats.MinHp = ModVida(UserList(UserIndex).raza).Inicial(UserList(UserIndex).clase)

    Dim MiInt As Integer
    MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
    If MiInt = 1 Then MiInt = 2
    
    UserList(UserIndex).Stats.MaxSta = 20 * MiInt
    UserList(UserIndex).Stats.MinSta = 20 * MiInt
    
    
    UserList(UserIndex).Stats.MaxAGU = 100
    UserList(UserIndex).Stats.MinAGU = 100
    
    UserList(UserIndex).Stats.MaxHam = 100
    UserList(UserIndex).Stats.MinHam = 100

    UserList(UserIndex).flags.ScrollExp = 1
    UserList(UserIndex).flags.ScrollOro = 1
    
    '<-----------------MANA----------------------->
    If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
        MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
        UserList(UserIndex).Stats.MaxMAN = MiInt
        UserList(UserIndex).Stats.MinMAN = MiInt
    ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
        Or UserClase = eClass.Bard Or UserClase = eClass.Paladin Or UserClase = eClass.Assasin Then
            UserList(UserIndex).Stats.MaxMAN = 50
            UserList(UserIndex).Stats.MinMAN = 50
    End If

    UserList(UserIndex).flags.VecesQueMoriste = 0
    UserList(UserIndex).flags.Montado = 0

    UserList(UserIndex).Stats.MaxHit = 2
    UserList(UserIndex).Stats.MinHIT = 1
    
    UserList(UserIndex).Stats.GLD = 0
    
    UserList(UserIndex).Stats.Exp = 0
    UserList(UserIndex).Stats.ELU = 300
    UserList(UserIndex).Stats.ELV = 1
    
    
    Call RellenarInventario(UserIndex)

    #If ConUpTime Then
        UserList(UserIndex).LogOnTime = Now
        UserList(UserIndex).UpTime = 0
    #End If
    
    'Valores Default de facciones al Activar nuevo usuario
    Call ResetFacciones(UserIndex)
    
    UserList(UserIndex).Faccion.Status = 1
    
    
    UserList(UserIndex).ChatCombate = 1
    UserList(UserIndex).ChatGlobal = 1
    
    
    'Resetamos CORREO
    UserList(UserIndex).Correo.CantCorreo = 0
    UserList(UserIndex).Correo.NoLeidos = 0
    'Resetamos CORREO
    
    If Not Database_Enabled Then
        Call GrabarNuevoPjEnCuentaCharfile(UserCuenta, name)
    End If
    
    UltimoChar = UCase$(name)
    
    Call SaveNewUser(UserIndex)
    Call ConnectUser(UserIndex, name, UserCuenta)
  
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Portal > 0 Then
        Dim Mapa As Integer
        Dim x As Byte
        Dim Y As Byte
        'Call SendData(SendTarget.ToMapButIndex, 0, PrepareMessageParticleFXToFloor(UserList(i).flags.PortalX, UserList(i).flags.PortalY, ParticulasIndex.Intermundia, 0))
        Mapa = UserList(UserIndex).flags.PortalM
        x = UserList(UserIndex).flags.PortalX
        Y = UserList(UserIndex).flags.PortalY
        'Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.Intermundia, 0))
        Call SendData(SendTarget.toMap, UserList(UserIndex).flags.PortalM, PrepareMessageParticleFXToFloor(x, Y, ParticulasIndex.TpVerde, 0))
        'Call SendData(SendTarget.toMap, 0, PrepareMessageParticleFXToFloor(X, Y, ParticulasIndex.Intermundia, 0))
        If MapData(Mapa, x, Y).TileExit.Map > 0 Then
            MapData(Mapa, x, Y).TileExit.Map = 0
            MapData(Mapa, x, Y).TileExit.x = 0
            MapData(Mapa, x, Y).TileExit.Y = 0
        End If
    End If

On Error GoTo Errhandler
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
    
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = UserIndex Then _
        Call modCentinela.CentinelaUserLogout
    
    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    'Empty buffer for reuse
    Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.length)
    
    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call MostrarNumUsers

        Call CloseUser(UserIndex)
        
    Else
        Call ResetUserSlot(UserIndex)
    End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

Exit Sub

Errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    
    Call ResetUserSlot(UserIndex)
    

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal UserIndex As Integer)
On Error GoTo Errhandler
    
    
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

    If UserIndex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers <> 0 Then NumUsers = NumUsers - 1
        Call MostrarNumUsers

        Call CloseUser(UserIndex)
    End If

    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    Call ResetUserSlot(UserIndex)

Exit Sub

Errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(UserIndex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo Errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(UserIndex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(UserIndex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

    If UserIndex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(UserIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call MostrarNumUsers
            NURestados = True
            Call CloseUser(UserIndex)
    End If
    
    Call ResetUserSlot(UserIndex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

Errhandler:
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call LogError("CLOSESOCKETERR: " & Err.description & " UI:" & UserIndex)
    
    If Not NURestados Then
        If UserList(UserIndex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
                Call MostrarNumUsers
            End If
            Call LogError("Cerre sin grabar a: " & UserList(UserIndex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(UserIndex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)

#If UsarQueSocket = 1 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call BorraSlotSock(UserList(UserIndex).ConnID)
    Call WSApiCloseSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    frmMain.Socket2(UserIndex).Cleanup
    Unload frmMain.Socket2(UserIndex)
    UserList(UserIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

#End If
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    Ret = WsApiEnviar(UserIndex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. userindex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************
    
    If frmMain.Socket2(UserIndex).Write(Datos, Len(Datos)) < 0 Then
        If frmMain.Socket2(UserIndex).LastError = WSAEWOULDBLOCK Then
            ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
            Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Datos)
        Else
            'Close the socket avoiding any critical error
            Call Cerrar_Usuario(UserIndex)
        End If
    End If
#ElseIf UsarQueSocket = 2 Then '**********************************************

    'Return value for this Socket:
    '--0) OK
    '--1) WSAEWOULDBLOCK
    '--2) ERROR
    
    Dim Ret As Long

    Ret = frmMain.Serv.Enviar(.ConnID, Datos, Len(Datos))
            
    If Ret = 1 Then
        ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
        Call .outgoingData.WriteASCIIStringFixed(Datos)
    ElseIf Ret = 2 Then
        'Close socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If
    

#ElseIf UsarQueSocket = 3 Then
    'THIS SOCKET DOESN`T USE THE BYTE QUEUE CLASS
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(UserIndex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un userIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(UserIndex).ConnID, Datos, Len(Datos)) = 2 Then Call CloseSocket(UserIndex)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For x = UserList(Index).Pos.x - MinXBorder + 1 To UserList(Index).Pos.x + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, x, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next x
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim x As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
            If x > 0 And Y > 0 And x < 101 And Y < 101 Then
                If MapData(Pos.Map, x, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next x
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim x As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1
            If MapData(Pos.Map, x, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next x
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal UserIndex As Integer) As Boolean

ValidateChr = UserList(UserIndex).Char.Head <> 0 _
                And UserList(UserIndex).Char.Body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Function EntrarCuenta(ByVal UserIndex As Integer, CuentaEmail As String, CuentaPassword As String, MacAddress As String, ByVal HDserial As Long) As Boolean

    If CheckMAC(MacAddress) Then
        Call WriteShowMessageBox(UserIndex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0001")
        Exit Function
    End If
    
    If CheckHD(HDserial) Then
        Call WriteShowMessageBox(UserIndex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0002")
        Exit Function
    End If

    If Not CheckMailString(CuentaEmail) Then
        Call WriteShowMessageBox(UserIndex, "Email inválido.")
        Exit Function
    End If
    
    
    If CuentaExiste(CuentaEmail) Then
        If Not ObtenerBaneo(CuentaEmail) Then
            If PasswordValida(CuentaEmail, SDesencriptar(CuentaPassword)) Then
                If ObtenerValidacion(CuentaEmail) Then
                    If Database_Enabled Then
                        UserList(UserIndex).AccountID = LoadCuentaDatabase(CuentaEmail, MacAddress, HDserial, UserList(UserIndex).ip)
                    Else
                        Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "MacAdress", MacAddress)
                        Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "HDserial", HDserial)
                        Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "UltimoAcceso", Date & " " & Time)
                        Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "UltimaIP", UserList(UserIndex).ip)
                    End If
                    
                    UserList(UserIndex).Cuenta = CuentaEmail
                    
                    EntrarCuenta = True
                Else
                    Call WriteShowMessageBox(UserIndex, "¡La cuenta no ha sido validada aún!")
                End If
            Else
               Call WriteShowMessageBox(UserIndex, "Contraseña inválida.")
            End If
        Else
            Call WriteShowMessageBox(UserIndex, "La cuenta se encuentra baneada debido a: " & ObtenerMotivoBaneo(CuentaEmail) & ". Esta decisión fue tomada por: " & ObtenerQuienBaneo(CuentaEmail) & ".")
        End If
    Else
       Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")
    End If
    
End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef name As String, ByRef UserCuenta As String)

On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim n As Integer
        Dim tStr As String
        
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .name & " ha intentado loguear a " & name & " desde la IP " & .ip)
            
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            
            Exit Sub
        End If
        
        'Reseteamos los FLAGS
        .flags.Escondido = 0
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Char.FX = 0
        
        'Controlamos no pasar el maximo de usuarios
        If NumUsers >= MaxUsers Then
            Call WriteShowMessageBox(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call FlushBuffer(UserIndex)
            'Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        '¿Este IP ya esta conectado?
        'If AllowMultiLogins = 0 Then
        '    If CheckForSameIP(UserIndex, .ip) = True Then
        '       ' Call WriteErrorMsg(UserIndex, "No es posible usar mas de un personaje al mismo tiempo.")
        '        Call WriteShowMessageBox(UserIndex, "No es posible usar mas de un personaje al mismo tiempo.")
        '        Call FlushBuffer(UserIndex)
        '       ' Call CloseSocket(UserIndex)
        '        Exit Sub
        '    End If
        'End If
        
        '¿Existe el personaje?
        'If Not FileExist(CharPath & UCase$(name) & ".chr", vbNormal) Then
        '    Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        '    Call FlushBuffer(UserIndex)
        '    Call CloseSocket(UserIndex)
        '    Exit Sub
        'End If

        'Reseteamos los privilegios
        .flags.Privilegios = 0
        
        'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
        If EsAdmin(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
        '    Call LogGM(name, "Se conecto con ip:" & .ip)
        ElseIf EsDios(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
        '    Call LogGM(name, "Se conecto con ip:" & .ip)
        ElseIf EsSemiDios(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
            Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
        '    Call LogGM(name, "Se conecto con ip:" & .ip)
        ElseIf EsConsejero(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
        '    Call LogGM(name, "Se conecto con ip:" & .ip)
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.user
            .flags.AdminPerseguible = True
        End If
        'Add RM flag if needed
        If EsRolesMaster(name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
            Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
        End If
    
        'If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
        '    If ObtenerLogeada(UCase$(UserCuenta)) = 1 Then
                ' Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                 'Call WriteShowMessageBox(UserIndex, "Solo se puede conectar un personaje por cuenta.")
                ' Call FlushBuffer(UserIndex)
                 'Call CloseSocket(UserIndex)
                 'Exit Sub
        '    End If
        'End If
        
        If ServerSoloGMs > 0 Then
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
               ' Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call FlushBuffer(UserIndex)
                'Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
        
        
        '¿Ya esta conectado el personaje?
        If CheckForSameName(name) Then
            If UserList(NameIndex(name)).Counters.Saliendo Then
                Call WriteShowMessageBox(UserIndex, "El usuario está saliendo.")
            Else
                Call WriteShowMessageBox(UserIndex, "Perdon, un usuario con el mismo nombre se há logoeado.")
            End If
            Call FlushBuffer(UserIndex)
           ' Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
        'Donador
        If DonadorCheck(UserCuenta) Then
            Dim LoopC As Integer
            For LoopC = 1 To Donadores.Count
                If UCase$(Donadores(LoopC).name) = UCase$(UserCuenta) Then
                    .donador.activo = 1
                    .donador.FechaExpiracion = Donadores(LoopC).FechaExpiracion
                    Exit For
                End If
            Next LoopC
        End If
        
        ' Seteamos el nombre
        .name = name
        
        ' Cargamos el personaje
        Call LoadUser(UserIndex)
        
        If Not ValidateChr(UserIndex) Then
            Call WriteShowMessageBox(UserIndex, "Error en el personaje. Comuniquese con el staff.")
            'Call CloseSocket(UserIndex)
            Exit Sub
        End If
    
        If UCase$(.Cuenta) <> UCase$(UserCuenta) Then
            Call WriteShowMessageBox(UserIndex, "El personaje no corresponde a su cuenta.")
           ' Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 And .Invent.NudilloSlot = 0 And .Invent.HerramientaEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        
        'If (.flags.Muerto = 0) Then
        '    .flags.SeguroResu = False
        '    Call WritePartySafeOff(UserIndex)
        'Else
            .flags.SeguroParty = True
        '    Call WritePartySafeOn(UserIndex)
        'End If
        
        .flags.SeguroClan = True
        
        
         .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
        
        Call WriteInventoryUnlockSlots(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)
        
        
        
        If .Correo.NoLeidos > 0 Then
            Call WriteCorreoPicOn(UserIndex)
        End If
        
        
        
        If .flags.Paralizado Then
            Call WriteParalizeOK(UserIndex)
        End If
        
        
        If .flags.Inmovilizado Then
            Call WriteInmovilizaOK(UserIndex)
        End If
        
        ''
        'TODO : Feo, esto tiene que ser parche cliente
        If .flags.Estupidez = 0 Then
            Call WriteDumbNoMore(UserIndex)
        End If
        
        'Posicion de comienzo
        If .Pos.Map = 0 Then
            .Pos.Map = 37
            .Pos.x = 76
            .Pos.Y = 82
        Else
            If Not MapaValido(.Pos.Map) Then
                Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
                Call FlushBuffer(UserIndex)
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
        End If
        
        If MapData(.Pos.Map, .Pos.x, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.x, .Pos.Y).NpcIndex <> 0 Then
            Call WarpToLegalPos(UserIndex, .Pos.Map, .Pos.x, .Pos.Y)
        End If
        
        'If in the water, and has a boat, equip it!
        If .Invent.BarcoObjIndex > 0 And HayAgua(.Pos.Map, .Pos.x, .Pos.Y) Then
            Dim Barco As ObjData
            Barco = ObjData(.Invent.BarcoObjIndex)
            If Barco.Ropaje <> iTraje Then
            .Char.Head = 0
            End If
            If .flags.Muerto = 0 Then
                '(Nacho)
                If .Faccion.ArmadaReal = 1 Then
                        If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCiuda
                        If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCiuda
                        If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCiuda
                        If Barco.Ropaje = iTraje Then .Char.Body = iTraje
                ElseIf .Faccion.FuerzasCaos = 1 Then
                        If Barco.Ropaje = iBarca Then .Char.Body = iBarcaPk
                        If Barco.Ropaje = iGalera Then .Char.Body = iGaleraPk
                        If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonPk
                        If Barco.Ropaje = iTraje Then .Char.Body = iTraje
                Else
                        If Barco.Ropaje = iBarca Then .Char.Body = iBarca
                        If Barco.Ropaje = iGalera Then .Char.Body = iGalera
                        If Barco.Ropaje = iGaleon Then .Char.Body = iGaleon
                        If Barco.Ropaje = iTraje Then .Char.Body = iTraje
                End If
            Else
                .Char.Body = iFragataFantasmal
            End If
            
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .flags.Navegando = 1
            
            .Char.speeding = Barco.Velocidad
            
            
            If Barco.Ropaje = iTraje Then
                Call WriteNadarToggle(UserIndex, True)
            Else
                Call WriteNadarToggle(UserIndex, False)
            End If
            
            Call WriteVelocidadToggle(UserIndex)
        End If

        Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        .flags.NecesitaOxigeno = RequiereOxigeno(.Pos.Map)
        
        Call WriteChangeMap(UserIndex, .Pos.Map) 'Carga el mapa
        Call WriteHora(UserIndex)
        
        'If .flags.Privilegios <> PlayerType.user And .flags.Privilegios <> (PlayerType.user Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.user Or PlayerType.RoyalCouncil) And .flags.Privilegios <> (PlayerType.user Or PlayerType.Admin) And .flags.Privilegios <> (PlayerType.user Or PlayerType.Dios) Then
           ' .flags.ChatColor = RGB(2, 161, 38)
        'ElseIf .flags.Privilegios = (PlayerType.user Or PlayerType.RoyalCouncil) Then
           ' .flags.ChatColor = RGB(0, 255, 255)
        If .flags.Privilegios = PlayerType.Admin Then
            .flags.ChatColor = RGB(217, 164, 32)
        ElseIf .flags.Privilegios = PlayerType.Dios Then
            .flags.ChatColor = RGB(217, 164, 32)
        ElseIf .flags.Privilegios = PlayerType.SemiDios Then
            .flags.ChatColor = RGB(2, 161, 38)
        ElseIf .flags.Privilegios = PlayerType.Consejero Then
            .flags.ChatColor = RGB(2, 161, 38)
        Else
            .flags.ChatColor = vbWhite
        End If
        
        
        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
        
        
        If .flags.Navegando = 0 Then
            If .flags.Muerto = 0 Then
                .Char.speeding = VelocidadNormal
            Else
                .Char.speeding = VelocidadMuerto
            End If
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        
        'Crea  el personaje del usuario
        Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.x, .Pos.Y, 1)

        Call WriteUserCharIndexInServer(UserIndex)

        ''[/el oso]
        
        Call WriteUpdateUserStats(UserIndex)
        
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        Call SendMOTD(UserIndex)

        If EnPausa Then
            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
        End If
        
        Call SetUserLogged(UserIndex)
        
        'Actualiza el Num de usuarios
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
        
        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
        If .Stats.SkillPts > 0 Then
            Call WriteSendSkills(UserIndex)
            Call WriteLevelUp(UserIndex, .Stats.SkillPts)
        End If
        
        
        If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
        
        If NumUsers > recordusuarios Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaniamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
            recordusuarios = NumUsers
            Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
        End If
        
        
        
        Call WriteFYA(UserIndex)
        Call WriteBindKeys(UserIndex)
        
        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(UserIndex)
        End If
        
        
        
        If .flags.Montado = 1 Then
            .Char.speeding = VelocidadMontura
            Call WriteEquiteToggle(UserIndex)
            'Debug.Print "Montado:" & .Char.speeding
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        End If
        
        If .flags.Muerto = 1 Then
            .Char.speeding = VelocidadMuerto
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        End If
        
        
        If .GuildIndex > 0 Then
            'welcome to the show baby...
            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
        
        
        tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
    
        If LenB(tStr) <> 0 Then
            Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)

        End If
        
        .flags.SolicitudPendienteDe = 0

        If Lloviendo Then
            Call WriteRainToggle(UserIndex)
        End If
        
        If ServidorNublado Then
            Call WriteNubesToggle(UserIndex)
        End If

        Call WriteLoggedMessage(UserIndex)
        
        If .Stats.ELV = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Bienvenido a las tierras de AO20! ¡" & .name & " que tengas buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
        ElseIf .Stats.ELV < 14 Then
            Call WriteConsoleMsg(UserIndex, "¡Bienvenido de nuevo " & .name & "! Actualmente estas en el nivel " & .Stats.ELV & " en " & DarNameMapa(.Pos.Map) & ", ¡buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
        End If

        If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
            Call WriteSafeModeOff(UserIndex)
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call WriteSafeModeOn(UserIndex)
        End If
        
        'Call modGuilds.SendGuildNews(UserIndex)
        
        If .MENSAJEINFORMACION <> vbNullString Then
            Call WriteConsoleMsg(UserIndex, .MENSAJEINFORMACION, FontTypeNames.FONTTYPE_CENTINELA)
            .MENSAJEINFORMACION = vbNullString
        End If

        tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
        
        If LenB(tStr) <> 0 Then
            Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
        End If

        If EventoActivo Then
            Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)
        End If
        
        Call WriteContadores(UserIndex)
        Call WriteOxigeno(UserIndex)

        'Call SendData(UserIndex, UserIndex, PrepareMessageParticleFXToFloor(.Pos.x, .Pos.y, 209, 10))
        'Call SendData(UserIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 209, 40, False))
        
        'Load the user statistics
        'Call Statistics.UserConnected(UserIndex)

        Call MostrarNumUsers
        'Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageParticleFXToFloor(.Pos.X, .Pos.y, ParticulasIndex.LogeoLevel1, 400))
        'Call SaveUser(UserIndex, CharPath & UCase$(.name) & ".chr")
        
       ' n = FreeFile
       ' Open App.Path & "\logs\numusers.log" For Output As n
        'Print #n, NumUsers
       ' Close #n

    End With
    
    Exit Sub
    
Errhandler:
    Call WriteShowMessageBox(UserIndex, "El personaje contiene un error, comuniquese con un miembro del staff.")
    Call FlushBuffer(UserIndex)
    
    'N = FreeFile
    'Log
    'Open App.Path & "\logs\Connect.log" For Append Shared As #N
    'Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
    'Close #N

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
    Dim j As Long
    For j = 1 To MaxLines
    Call WriteConsoleMsg(UserIndex, MOTD(j).texto, FontTypeNames.FONTTYPE_EXP)
    Next j
    
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .ArmadaReal = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .Status = 0
        .FuerzasCaos = 0
        .FechaIngreso = ""
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Inmovilizado = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Maldicion = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        'Ladder
        .Incineracion = 0
        'Ladder
        .ScrollExperiencia = 0
        .ScrollOro = 0
        .Oxigeno = 0
        .TiempoParaSubastar = 0
        .TimerPerteneceNpc = 0
        .TimerPuedeSerAtacado = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************

    With UserList(UserIndex).Char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
        .Arma_Aura = ""
        .Body_Aura = ""
        .Head_Aura = ""
        .Otra_Aura = ""
        .Escudo_Aura = ""
        .ParticulaFx = 0
        .speeding = VelocidadCero
    End With
    

End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'Agregue que se resetee el maná
'*************************************************
Dim LoopC As Integer
    With UserList(UserIndex)
        .name = vbNullString
        .Cuenta = vbNullString
        .Id = -1
        .AccountID = -1
        .modName = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.x = 0
        .Pos.Y = 0
        .ip = vbNullString
        .clase = 0
        .email = vbNullString
        .genero = 0
        .Hogar = 0
        .raza = 0
        .EmpoCont = 0
        
        'Ladder     Reseteo de Correos
        .Correo.CantCorreo = 0
        .Correo.NoLeidos = 0
        
        For LoopC = 1 To MAX_CORREOS_SLOTS
                    .Correo.Mensaje(LoopC).Remitente = ""
                    .Correo.Mensaje(LoopC).Mensaje = ""
                    .Correo.Mensaje(LoopC).Item = 0
                    .Correo.Mensaje(LoopC).ItemCount = 0
                    .Correo.Mensaje(LoopC).Fecha = ""
                    .Correo.Mensaje(LoopC).Leido = 0
        Next LoopC
        'Ladder     Reseteo de Correos
        

        

        
        
        
        With .Stats
            .InventLevel = 0
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .MaxMAN = 0
            .MinMAN = 0
            
        End With
        
    End With
End Sub


Sub ResetGuildInfo(ByVal UserIndex As Integer)
    If UserList(UserIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
        UserList(UserIndex).EscucheClan = 0
    End If
    If UserList(UserIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)
    End If
    UserList(UserIndex).GuildIndex = 0
    
    
    
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************

    With UserList(UserIndex).flags
        .LevelBackup = 0
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .ScrollExp = 1
        .ScrollOro = 1
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .Ahogandose = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Escribiendo = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        'Ladder
        .VecesQueMoriste = 0
        .MinutosRestantes = 0
        .SegundosPasados = 0
        .RetoA = 0
        .SolicitudPendienteDe = 0
        .CarroMineria = 0
        .DañoMagico = 0
        .Montado = 0
        .Incinerado = 0
        .Casado = 0
        .Pareja = ""
        .Candidato = 0
        .UsandoMacro = False
        .pregunta = 0
        'Ladder
        .BattleModo = 0
        .ResistenciaMagica = 0

        .Subastando = False
        .Paraliza = 0
        .Envenena = 0
        .NoPalabrasMagicas = 0
        .NoMagiaEfeceto = 0
        .incinera = 0
        .Estupidiza = 0
        .GolpeCertero = 0
        .PendienteDelExperto = 0
        .CarroMineria = 0
        .PendienteDelSacrificio = 0
        .AnilloOcultismo = 0
        .RegeneracionMana = 0
        .RegeneracionHP = 0
        .RegeneracionSta = 0
        .NecesitaOxigeno = False
        .LastCrimMatado = ""
        .LastCiudMatado = ""
        
    End With
End Sub
Sub ResetAccionesPendientes(ByVal UserIndex As Integer)
'*************************************************
'*************************************************
    With UserList(UserIndex).accion
        .AccionPendiente = False
        .HechizoPendiente = 0
        .RunaObj = 0
        .Particula = 0
        .TipoAccion = 0
        .ObjSlot = 0
    End With
End Sub
Sub ResetDonadorFlag(ByVal UserIndex As Integer)
'*************************************************
'*************************************************
    With UserList(UserIndex).donador
        .activo = 0
        .CreditoDonador = 0
        .FechaExpiracion = 0
    End With
End Sub
Sub ResetUserSpells(ByVal UserIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
       ' UserList(UserIndex).Stats.UserHechizosInterval(LoopC) = 0
    Next LoopC
End Sub
Sub ResetUserSkills(ByVal UserIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To NUMSKILLS
        UserList(UserIndex).Stats.UserSkills(LoopC) = 0
    Next LoopC
End Sub


Sub ResetUserBanco(ByVal UserIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)



UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1


If UserList(UserIndex).Grupo.Lider = UserIndex Then
    Call FinalizarGrupo(UserIndex)
End If

If UserList(UserIndex).Grupo.EnGrupo Then
    Call SalirDeGrupoForzado(UserIndex)
End If




UserList(UserIndex).Grupo.CantidadMiembros = 0
UserList(UserIndex).Grupo.EnGrupo = False
UserList(UserIndex).Grupo.Lider = 0
UserList(UserIndex).Grupo.PropuestaDe = 0
UserList(UserIndex).Grupo.Miembros(6) = 0
UserList(UserIndex).Grupo.Miembros(1) = 0
UserList(UserIndex).Grupo.Miembros(2) = 0
UserList(UserIndex).Grupo.Miembros(3) = 0
UserList(UserIndex).Grupo.Miembros(4) = 0
UserList(UserIndex).Grupo.Miembros(5) = 0


Call ResetQuestStats(UserIndex)
Call ResetGuildInfo(UserIndex)
Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call ResetAccionesPendientes(UserIndex)
Call ResetDonadorFlag(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
'Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
Call ResetUserSkills(UserIndex)
With UserList(UserIndex).ComUsu
    .Acepto = False
    .cant = 0
    .DestNick = vbNullString
    .DestUsu = 0
    .Objeto = 0
End With

End Sub

Sub CloseUser(ByVal UserIndex As Integer)
    'Call LogTarea("CloseUser " & UserIndex)
    On Error GoTo Errhandler
    
    Dim n As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim LoopC As Integer
    Dim Map As Integer
    Dim name As String
    Dim raza As eRaza
    Dim clase As eClass
    Dim i As Integer
    
    Dim aN As Integer
    
    
    Dim errordesc As String
    
    
    errordesc = "ERROR AL SETEAR NPC"
    aN = UserList(UserIndex).flags.AtacadoPorNpc
    If aN > 0 Then
          Npclist(aN).Movement = Npclist(aN).flags.OldMovement
          Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
          Npclist(aN).flags.AttackedBy = vbNullString
    End If
    aN = UserList(UserIndex).flags.NPCAtacado
    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString
        End If
    End If
    UserList(UserIndex).flags.AtacadoPorNpc = 0
    UserList(UserIndex).flags.NPCAtacado = 0
    
    If UserList(UserIndex).flags.Montado > 0 Then
        Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)
    End If
    
    'CHECK:: ACA SE GUARDAN UN MONTON DE COSAS QUE NO SE OCUPAN PARA NADA :S
    Map = UserList(UserIndex).Pos.Map
    x = UserList(UserIndex).Pos.x
    Y = UserList(UserIndex).Pos.Y
    name = UCase$(UserList(UserIndex).name)
    raza = UserList(UserIndex).raza
    clase = UserList(UserIndex).clase
    
    
    errordesc = "ERROR AL ENVIAR PARTICULA"
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.loops = 0
    UserList(UserIndex).Char.ParticulaFx = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 0, 0, True))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    
    UserList(UserIndex).flags.UserLogged = False
    UserList(UserIndex).Counters.Saliendo = False
    
    
    errordesc = "ERROR AL ENVIAR INVI"
    
    'Le devolvemos el body y head originales
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)
    
    
    errordesc = "ERROR AL CANCELAR SUBASTA"
    If UserList(UserIndex).flags.Subastando = True Then
        Call CancelarSubasta
    End If
    
    
    errordesc = "ERROR AL BORRAR INDEX DE TORNEO"
    If UserList(UserIndex).flags.EnTorneo = True Then
        Call BorrarIndexInTorneo(UserIndex)
        UserList(UserIndex).flags.EnTorneo = False
    End If
    
    
    'Save statistics
    'Call Statistics.UserDisconnected(UserIndex)
    
    ' Grabamos el personaje del usuario
    
    
    errordesc = "ERROR AL GRABAR PJ"
    If UserList(UserIndex).flags.BattleModo = 0 Then
        Call SaveUser(UserIndex, True)
    Else
        'Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)
        Call SaveBattlePoints(UserIndex)
    End If

    errordesc = "ERROR AL DESCONTAR USER DE MAPA"
    If MapInfo(Map).NumUsers > 0 Then
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
    End If

    errordesc = "ERROR AL ERASEUSERCHAR"
    'Borrar el personaje
    If UserList(UserIndex).Char.CharIndex > 0 Then
        Call EraseUserChar(UserIndex, True)
    End If
    
    errordesc = "ERROR Update Map Users"
    'Update Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    
    If MapInfo(Map).NumUsers < 0 Then
        MapInfo(Map).NumUsers = 0
    End If
    
    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    'If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)
    errordesc = "ERROR AL RESETSLOP Name:" & UserList(UserIndex).name & " cuenta:" & UserList(UserIndex).Cuenta
    Call ResetUserSlot(UserIndex)
    
    Exit Sub
    
Errhandler:
    Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.description & "Detalle:" & errordesc)

End Sub

Sub ReloadSokcet()
On Error GoTo Errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
Errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub


Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.user Then
            Call CloseSocket(LoopC)
        End If
    End If
Next LoopC

End Sub
