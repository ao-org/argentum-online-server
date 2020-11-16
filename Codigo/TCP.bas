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
    Public Const INVALID_HANDLE     As Integer = -1

    Public Const CONTROL_ERRIGNORE  As Integer = 0

    Public Const CONTROL_ERRDISPLAY As Integer = 1

    ' SocietWrench Control Actions
    Public Const SOCKET_OPEN        As Integer = 1

    Public Const SOCKET_CONNECT     As Integer = 2

    Public Const SOCKET_LISTEN      As Integer = 3

    Public Const SOCKET_ACCEPT      As Integer = 4

    Public Const SOCKET_CANCEL      As Integer = 5

    Public Const SOCKET_FLUSH       As Integer = 6

    Public Const SOCKET_CLOSE       As Integer = 7

    Public Const SOCKET_DISCONNECT  As Integer = 7

    Public Const SOCKET_ABORT       As Integer = 8

    ' SocketWrench Control States
    Public Const SOCKET_NONE        As Integer = 0

    Public Const SOCKET_IDLE        As Integer = 1

    Public Const SOCKET_LISTENING   As Integer = 2

    Public Const SOCKET_CONNECTING  As Integer = 3

    Public Const SOCKET_ACCEPTING   As Integer = 4

    Public Const SOCKET_RECEIVING   As Integer = 5

    Public Const SOCKET_SENDING     As Integer = 6

    Public Const SOCKET_CLOSING     As Integer = 7

    ' Societ Address Families
    Public Const AF_UNSPEC          As Integer = 0

    Public Const AF_UNIX            As Integer = 1

    Public Const AF_INET            As Integer = 2

    ' Societ Types
    Public Const SOCK_STREAM        As Integer = 1

    Public Const SOCK_DGRAM         As Integer = 2

    Public Const SOCK_RAW           As Integer = 3

    Public Const SOCK_RDM           As Integer = 4

    Public Const SOCK_SEQPACKET     As Integer = 5

    ' Protocol Types
    Public Const IPPROTO_IP         As Integer = 0

    Public Const IPPROTO_ICMP       As Integer = 1

    Public Const IPPROTO_GGP        As Integer = 2

    Public Const IPPROTO_TCP        As Integer = 6

    Public Const IPPROTO_PUP        As Integer = 12

    Public Const IPPROTO_UDP        As Integer = 17

    Public Const IPPROTO_IDP        As Integer = 22

    Public Const IPPROTO_ND         As Integer = 77

    Public Const IPPROTO_RAW        As Integer = 255

    Public Const IPPROTO_MAX        As Integer = 256

    ' Network Addpesses
    Public Const INADDR_ANY         As String = "0.0.0.0"

    Public Const INADDR_LOOPBACK    As String = "127.0.0.1"

    Public Const INADDR_NONE        As String = "255.055.255.255"

    ' Shutdown Values
    Public Const SOCKET_READ        As Integer = 0

    Public Const SOCKET_WRITE       As Integer = 1

    Public Const SOCKET_READWRITE   As Integer = 2

    ' SocketWrench Error Pesponse
    Public Const SOCKET_ERRIGNORE   As Integer = 0

    Public Const SOCKET_ERRDISPLAY  As Integer = 1

    ' SocketWrench Error Codes
    Public Const WSABASEERR         As Integer = 24000

    Public Const WSAEINTR           As Integer = 24004

    Public Const WSAEBADF           As Integer = 24009

    Public Const WSAEACCES          As Integer = 24013

    Public Const WSAEFAULT          As Integer = 24014

    Public Const WSAEINVAL          As Integer = 24022

    Public Const WSAEMFILE          As Integer = 24024

    Public Const WSAEWOULDBLOCK     As Integer = 24035

    Public Const WSAEINPROGRESS     As Integer = 24036

    Public Const WSAEALREADY        As Integer = 24037

    Public Const WSAENOTSOCK        As Integer = 24038

    Public Const WSAEDESTADDRREQ    As Integer = 24039

    Public Const WSAEMSGSIZE        As Integer = 24040

    Public Const WSAEPROTOTYPE      As Integer = 24041

    Public Const WSAENOPROTOOPT     As Integer = 24042

    Public Const WSAEPROTONOSUPPORT As Integer = 24043

    Public Const WSAESOCKTNOSUPPORT As Integer = 24044

    Public Const WSAEOPNOTSUPP      As Integer = 24045

    Public Const WSAEPFNOSUPPORT    As Integer = 24046

    Public Const WSAEAFNOSUPPORT    As Integer = 24047

    Public Const WSAEADDRINUSE      As Integer = 24048

    Public Const WSAEADDRNOTAVAIL   As Integer = 24049

    Public Const WSAENETDOWN        As Integer = 24050

    Public Const WSAENETUNREACH     As Integer = 24051

    Public Const WSAENETRESET       As Integer = 24052

    Public Const WSAECONNABORTED    As Integer = 24053

    Public Const WSAECONNRESET      As Integer = 24054

    Public Const WSAENOBUFS         As Integer = 24055

    Public Const WSAEISCONN         As Integer = 24056

    Public Const WSAENOTCONN        As Integer = 24057

    Public Const WSAESHUTDOWN       As Integer = 24058

    Public Const WSAETOOMANYREFS    As Integer = 24059

    Public Const WSAETIMEDOUT       As Integer = 24060

    Public Const WSAECONNREFUSED    As Integer = 24061

    Public Const WSAELOOP           As Integer = 24062

    Public Const WSAENAMETOOLONG    As Integer = 24063

    Public Const WSAEHOSTDOWN       As Integer = 24064

    Public Const WSAEHOSTUNREACH    As Integer = 24065

    Public Const WSAENOTEMPTY       As Integer = 24066

    Public Const WSAEPROCLIM        As Integer = 24067

    Public Const WSAEUSERS          As Integer = 24068

    Public Const WSAEDQUOT          As Integer = 24069

    Public Const WSAESTALE          As Integer = 24070

    Public Const WSAEREMOTE         As Integer = 24071

    Public Const WSASYSNOTREADY     As Integer = 24091

    Public Const WSAVERNOTSUPPORTED As Integer = 24092

    Public Const WSANOTINITIALISED  As Integer = 24093

    Public Const WSAHOST_NOT_FOUND  As Integer = 25001

    Public Const WSATRY_AGAIN       As Integer = 25002

    Public Const WSANO_RECOVERY     As Integer = 25003

    Public Const WSANO_DATA         As Integer = 25004

    Public Const WSANO_ADDRESS      As Integer = 2500

#End If

Sub DarCuerpo(ByVal UserIndex As Integer)
        
        On Error GoTo DarCuerpo_Err
        

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 14/03/2007
        'Elije una cabeza para el usuario y le da un body
        '*************************************************
        Dim NewBody    As Integer

        Dim UserRaza   As Byte

        Dim UserGenero As Byte

100     UserGenero = UserList(UserIndex).genero
102     UserRaza = UserList(UserIndex).raza

104     Select Case UserGenero

            Case eGenero.Hombre

106             Select Case UserRaza

                    Case eRaza.Humano
108                     NewBody = 1

110                 Case eRaza.Elfo
112                     NewBody = 2

114                 Case eRaza.Drow
116                     NewBody = 3

118                 Case eRaza.Enano
120                     NewBody = 300

122                 Case eRaza.Gnomo
124                     NewBody = 300

126                 Case eRaza.Orco
128                     NewBody = 582

                End Select

130         Case eGenero.Mujer

132             Select Case UserRaza

                    Case eRaza.Humano
134                     NewBody = 1

136                 Case eRaza.Elfo
138                     NewBody = 2

140                 Case eRaza.Drow
142                     NewBody = 3

144                 Case eRaza.Gnomo
146                     NewBody = 300

148                 Case eRaza.Enano
150                     NewBody = 300

152                 Case eRaza.Orco
154                     NewBody = 581

                End Select

        End Select

156     UserList(UserIndex).Char.Body = NewBody

        
        Exit Sub

DarCuerpo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.DarCuerpo", Erl)
        Resume Next
        
End Sub

Sub AsignarAtributos(ByVal UserIndex As String)
        
        On Error GoTo AsignarAtributos_Err
        

100     Select Case UserList(UserIndex).raza

            Case eRaza.Humano
102             UserList(UserIndex).Stats.UserAtributos(1) = 19
104             UserList(UserIndex).Stats.UserAtributos(2) = 19
106             UserList(UserIndex).Stats.UserAtributos(3) = 19
108             UserList(UserIndex).Stats.UserAtributos(4) = 20

110         Case eRaza.Elfo
112             UserList(UserIndex).Stats.UserAtributos(1) = 18
114             UserList(UserIndex).Stats.UserAtributos(2) = 20
116             UserList(UserIndex).Stats.UserAtributos(3) = 21
118             UserList(UserIndex).Stats.UserAtributos(4) = 18

120         Case eRaza.Drow
122             UserList(UserIndex).Stats.UserAtributos(1) = 20
124             UserList(UserIndex).Stats.UserAtributos(2) = 18
126             UserList(UserIndex).Stats.UserAtributos(3) = 20
128             UserList(UserIndex).Stats.UserAtributos(4) = 19

130         Case eRaza.Gnomo
132             UserList(UserIndex).Stats.UserAtributos(1) = 13
134             UserList(UserIndex).Stats.UserAtributos(2) = 21
136             UserList(UserIndex).Stats.UserAtributos(3) = 22
138             UserList(UserIndex).Stats.UserAtributos(4) = 17

140         Case eRaza.Enano
142             UserList(UserIndex).Stats.UserAtributos(1) = 21
144             UserList(UserIndex).Stats.UserAtributos(2) = 17
146             UserList(UserIndex).Stats.UserAtributos(3) = 12
148             UserList(UserIndex).Stats.UserAtributos(4) = 22

150         Case eRaza.Orco
152             UserList(UserIndex).Stats.UserAtributos(1) = 23
154             UserList(UserIndex).Stats.UserAtributos(2) = 17
156             UserList(UserIndex).Stats.UserAtributos(3) = 12
158             UserList(UserIndex).Stats.UserAtributos(4) = 21

        End Select

        
        Exit Sub

AsignarAtributos_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.AsignarAtributos", Erl)
        Resume Next
        
End Sub

Sub RellenarInventario(ByVal UserIndex As String)
        
        On Error GoTo RellenarInventario_Err
        

100     With UserList(UserIndex)
        
            Dim NumItems As Integer

102         NumItems = 1
    
            ' Todos reciben pociones rojas
104         .Invent.Object(NumItems).ObjIndex = 1616 'Pocion Roja
106         .Invent.Object(NumItems).Amount = 100
108         NumItems = NumItems + 1
        
            ' Magicas puras reciben más azules
110         Select Case .clase

                Case eClass.Mage, eClass.Druid
112                 .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
114                 .Invent.Object(NumItems).Amount = 100
116                 NumItems = NumItems + 1

            End Select
        
            ' Semi mágicas reciben menos
118         Select Case .clase

                Case eClass.Bard, eClass.Cleric, eClass.Paladin, eClass.Assasin
120                 .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
122                 .Invent.Object(NumItems).Amount = 50
124                 NumItems = NumItems + 1

            End Select

            ' Arma y hechizos
126         Select Case .clase

                Case eClass.Mage, eClass.Cleric, eClass.Druid, eClass.Bard
128                 .Stats.UserHechizos(1) = 1 ' Proyectil
130                 .Stats.UserHechizos(2) = 2 ' Saeta
132                 .Stats.UserHechizos(3) = 11 ' Curar Veneno
134                 .Stats.UserHechizos(4) = 12 ' Heridas Leves

136             Case eClass.Assasin, eClass.Paladin
138                 .Stats.UserHechizos(1) = 1 ' Proyectil
140                 .Stats.UserHechizos(2) = 2 ' Saeta
142                 .Stats.UserHechizos(3) = 11 ' Curar Veneno

            End Select
        
            ' Pociones amarillas y verdes
144         Select Case .clase

                Case eClass.Assasin, eClass.Bard, eClass.Cleric, eClass.Hunter, eClass.Paladin, eClass.Trabajador, eClass.Warrior
146                 .Invent.Object(NumItems).ObjIndex = 1618 ' Pocion Amarilla
148                 .Invent.Object(NumItems).Amount = 25
150                 NumItems = NumItems + 1

152                 .Invent.Object(NumItems).ObjIndex = 1619 ' Pocion Verde
154                 .Invent.Object(NumItems).Amount = 25
156                 NumItems = NumItems + 1

            End Select
            
            ' Poción violeta
158         .Invent.Object(NumItems).ObjIndex = 166 ' Pocion violeta
159         .Invent.Object(NumItems).Amount = 10
160         NumItems = NumItems + 1
        
            ' Equipo el arma
161         .Invent.Object(NumItems).ObjIndex = 460 ' Daga (Newbies)
162         .Invent.Object(NumItems).Amount = 1
163         .Invent.Object(NumItems).Equipped = 1
164         .Invent.WeaponEqpSlot = NumItems
166         .Invent.WeaponEqpObjIndex = .Invent.Object(NumItems).ObjIndex
168         .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
170         NumItems = NumItems + 1
        
            
            If .genero = eGenero.Hombre Then
                If .raza = Enano Or .raza = Gnomo Then
                    .Invent.Object(NumItems).ObjIndex = 466 'Vestimentas de Bajo (Newbies)
                Else
                
                    .Invent.Object(NumItems).ObjIndex = RandomNumber(463, 465) ' Vestimentas comunes (Newbies)
                End If
            Else
                If .raza = Enano Or .raza = Gnomo Then
                    .Invent.Object(NumItems).ObjIndex = 563 'Vestimentas de Baja (Newbies)
                Else
                    .Invent.Object(NumItems).ObjIndex = RandomNumber(1283, 1285) ' Vestimentas de Mujer (Newbies)
                End If
            End If
                        
174         .Invent.Object(NumItems).Amount = 1
176         .Invent.Object(NumItems).Equipped = 1
178         .Invent.ArmourEqpSlot = NumItems
180         .Invent.ArmourEqpObjIndex = .Invent.Object(NumItems).ObjIndex
182          NumItems = NumItems + 1

            ' Animación según raza
184
188          .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        
            ' Comida y bebida
190         .Invent.Object(NumItems).ObjIndex = 573 ' Manzana
192         .Invent.Object(NumItems).Amount = 100
194         NumItems = NumItems + 1

196         .Invent.Object(NumItems).ObjIndex = 572 ' Agua
198         .Invent.Object(NumItems).Amount = 100
200         NumItems = NumItems + 1

202         .Invent.Object(NumItems).ObjIndex = 200 ' Cofre Inicial - TODO: Remover
204         .Invent.Object(NumItems).Amount = 1
206         NumItems = NumItems + 1

            ' Seteo la cantidad de items
208         .Invent.NroItems = NumItems

        End With
   
        
        Exit Sub

RellenarInventario_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.RellenarInventario", Erl)
        Resume Next
        
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
        
        On Error GoTo AsciiValidos_Err
        

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
108             AsciiValidos = False
                Exit Function

            End If
    
110     Next i

112     AsciiValidos = True

        
        Exit Function

AsciiValidos_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.AsciiValidos", Erl)
        Resume Next
        
End Function

Function Numeric(ByVal cad As String) As Boolean
        
        On Error GoTo Numeric_Err
        

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If (car < 48 Or car > 57) Then
108             Numeric = False
                Exit Function

            End If
    
110     Next i

112     Numeric = True

        
        Exit Function

Numeric_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.Numeric", Erl)
        Resume Next
        
End Function

Function NombrePermitido(ByVal nombre As String) As Boolean
        
        On Error GoTo NombrePermitido_Err
        

        Dim i As Integer

100     For i = 1 To UBound(ForbidenNames)

102         If InStr(nombre, ForbidenNames(i)) Then
104             NombrePermitido = False
                Exit Function

            End If

106     Next i

108     NombrePermitido = True

        
        Exit Function

NombrePermitido_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.NombrePermitido", Erl)
        Resume Next
        
End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo ValidateSkills_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To NUMSKILLS

102         If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
                Exit Function

104             If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100

            End If

106     Next LoopC

108     ValidateSkills = True
    
        
        Exit Function

ValidateSkills_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ValidateSkills", Erl)
        Resume Next
        
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Head As Integer, ByRef UserCuenta As String)
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
        
        On Error GoTo ConnectNewUser_Err
        
    
100     If Not AsciiValidos(name) Or LenB(name) = 0 Then
102         Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Exit Sub

        End If
    
104     If UserList(UserIndex).flags.UserLogged Then
106         Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado crear a " & name & " desde la IP " & UserList(UserIndex).ip)
108         Call CloseSocketSL(UserIndex)
110         Call Cerrar_Usuario(UserIndex)
            Exit Sub

        End If
    
        Dim LoopC As Long
    
        '¿Existe el personaje?
112     If PersonajeExiste(name) Then
114         Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
            Exit Sub

        End If
    
        'Prevenimos algun bug con dados inválidos
116     If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then Exit Sub
    
118     UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
120     UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
122     UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
124     UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
        UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
    
126     UserList(UserIndex).flags.Muerto = 0
128     UserList(UserIndex).flags.Escondido = 0

130     UserList(UserIndex).flags.Casado = 0
132     UserList(UserIndex).flags.Pareja = ""

134     UserList(UserIndex).name = name
136     UserList(UserIndex).clase = UserClase
138     UserList(UserIndex).raza = UserRaza
    
140     UserList(UserIndex).Char.Head = Head
    
142     UserList(UserIndex).genero = UserSexo
144     UserList(UserIndex).Hogar = 1
    
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
146     UserList(UserIndex).Stats.SkillPts = 10
    
148     UserList(UserIndex).Char.heading = eHeading.SOUTH
    
150     Call DarCuerpo(UserIndex) 'Ladder REVISAR
    
152     UserList(UserIndex).OrigChar = UserList(UserIndex).Char

154     UserList(UserIndex).Char.WeaponAnim = NingunArma
156     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
158     UserList(UserIndex).Char.CascoAnim = NingunCasco

        'Call AsignarAtributos(UserIndex)

        Dim MiInt As Integer
    
160     MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
162     UserList(UserIndex).Stats.MaxHp = 15 + MiInt
164     UserList(UserIndex).Stats.MinHp = 15 + MiInt
    
166     MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)

168     If MiInt = 1 Then MiInt = 2
    
170     UserList(UserIndex).Stats.MaxSta = 20 * MiInt
172     UserList(UserIndex).Stats.MinSta = 20 * MiInt
    
174     UserList(UserIndex).Stats.MaxAGU = 100
176     UserList(UserIndex).Stats.MinAGU = 100
    
178     UserList(UserIndex).Stats.MaxHam = 100
180     UserList(UserIndex).Stats.MinHam = 100

182     UserList(UserIndex).flags.ScrollExp = 1
184     UserList(UserIndex).flags.ScrollOro = 1
    
        '<-----------------MANA----------------------->
186     If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
188         MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
190         UserList(UserIndex).Stats.MaxMAN = MiInt
192         UserList(UserIndex).Stats.MinMAN = MiInt
194     ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid Or UserClase = eClass.Bard Or UserClase = eClass.Paladin Or UserClase = eClass.Assasin Then
196         UserList(UserIndex).Stats.MaxMAN = 50
198         UserList(UserIndex).Stats.MinMAN = 50

        End If

200     UserList(UserIndex).flags.VecesQueMoriste = 0
202     UserList(UserIndex).flags.Montado = 0

204     UserList(UserIndex).Stats.MaxHit = 2
206     UserList(UserIndex).Stats.MinHIT = 1
    
208     UserList(UserIndex).Stats.GLD = 0
    
210     UserList(UserIndex).Stats.Exp = 0
212     UserList(UserIndex).Stats.ELU = 300
214     UserList(UserIndex).Stats.ELV = 1
    
216     Call RellenarInventario(UserIndex)

        #If ConUpTime Then
218         UserList(UserIndex).LogOnTime = Now
220         UserList(UserIndex).UpTime = 0
        #End If
    
        'Valores Default de facciones al Activar nuevo usuario
222     Call ResetFacciones(UserIndex)
    
224     UserList(UserIndex).Faccion.Status = 1
    
226     UserList(UserIndex).ChatCombate = 1
228     UserList(UserIndex).ChatGlobal = 1
    
        'Resetamos CORREO
230     UserList(UserIndex).Correo.CantCorreo = 0
232     UserList(UserIndex).Correo.NoLeidos = 0
        'Resetamos CORREO
    
234     UserList(UserIndex).Pos.Map = 37
236     UserList(UserIndex).Pos.x = 76
238     UserList(UserIndex).Pos.Y = 82
    
240     If Not Database_Enabled Then
242         Call GrabarNuevoPjEnCuentaCharfile(UserCuenta, name)
        End If
    
244     UltimoChar = UCase$(name)
    
246     Call SaveNewUser(UserIndex)
248     Call ConnectUser(UserIndex, name, UserCuenta)
        
        Exit Sub

ConnectNewUser_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ConnectNewUser", Erl)
        Resume Next
        
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

    Sub CloseSocket(ByVal UserIndex As Integer)
    
        Call FlushBuffer(UserIndex)

        If UserList(UserIndex).flags.Portal > 0 Then

            Dim Mapa As Integer

            Dim x    As Byte

            Dim Y    As Byte

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
        If Centinela.RevisandoUserIndex = UserIndex Then Call modCentinela.CentinelaUserLogout
    
        'mato los comercios seguros
        If UserList(UserIndex).ComUsu.DestUsu > 0 Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    

                End If

            End If

        End If
    
        'Empty buffer for reuse
        Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.length)
    
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseUser(UserIndex)
        
            If NumUsers > 0 Then NumUsers = NumUsers - 1
            Call MostrarNumUsers
        
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
        
        On Error GoTo CloseSocketSL_Err
        

        #If UsarQueSocket = 1 Then

100         If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
102             Call BorraSlotSock(UserList(UserIndex).ConnID)
104             Call WSApiCloseSocket(UserList(UserIndex).ConnID)
106             UserList(UserIndex).ConnIDValida = False

            End If

        #ElseIf UsarQueSocket = 0 Then

108         If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
110             frmMain.Socket2(UserIndex).Cleanup
112             Unload frmMain.Socket2(UserIndex)
114             UserList(UserIndex).ConnIDValida = False

            End If

        #ElseIf UsarQueSocket = 2 Then

116         If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
118             Call frmMain.Serv.CerrarSocket(UserList(UserIndex).ConnID)
120             UserList(UserIndex).ConnIDValida = False

            End If

        #End If

        
        Exit Sub

CloseSocketSL_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.CloseSocketSL", Erl)
        Resume Next
        
End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Sub EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: 09/11/20
        'Last Modified By: Jopi
        'Se agrega el paquete a la cola, para prevenir errores.
        '***************************************************
        
        On Error GoTo EnviarDatosASlot_Err
        

100     Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Datos)

        Exit Sub

ErrorHandler:
102     Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & UserIndex & "/" & UserList(UserIndex).ConnID & "/" & Datos)

        
        Exit Sub

EnviarDatosASlot_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EnviarDatosASlot", Erl)
        Resume Next
        
End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
        
        On Error GoTo EstaPCarea_Err
        

        Dim x As Integer, Y As Integer

100     For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
102         For x = UserList(Index).Pos.x - MinXBorder + 1 To UserList(Index).Pos.x + MinXBorder - 1

104             If MapData(UserList(Index).Pos.Map, x, Y).UserIndex = Index2 Then
106                 EstaPCarea = True
                    Exit Function

                End If
        
108         Next x
110     Next Y

112     EstaPCarea = False

        
        Exit Function

EstaPCarea_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EstaPCarea", Erl)
        Resume Next
        
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
        
        On Error GoTo HayPCarea_Err
        

        Dim x As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1

104             If x > 0 And Y > 0 And x < 101 And Y < 101 Then
106                 If MapData(Pos.Map, x, Y).UserIndex > 0 Then
108                     HayPCarea = True
                        Exit Function

                    End If

                End If

110         Next x
112     Next Y

114     HayPCarea = False

        
        Exit Function

HayPCarea_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.HayPCarea", Erl)
        Resume Next
        
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
        
        On Error GoTo HayOBJarea_Err
        

        Dim x As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For x = Pos.x - MinXBorder + 1 To Pos.x + MinXBorder - 1

104             If MapData(Pos.Map, x, Y).ObjInfo.ObjIndex = ObjIndex Then
106                 HayOBJarea = True
                    Exit Function

                End If
        
108         Next x
110     Next Y

112     HayOBJarea = False

        
        Exit Function

HayOBJarea_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.HayOBJarea", Erl)
        Resume Next
        
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo ValidateChr_Err
        

100     ValidateChr = UserList(UserIndex).Char.Head <> 0 And UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

        
        Exit Function

ValidateChr_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ValidateChr", Erl)
        Resume Next
        
End Function

Function EntrarCuenta(ByVal UserIndex As Integer, CuentaEmail As String, CuentaPassword As String, MacAddress As String, ByVal HDserial As Long) As Boolean
        
        On Error GoTo EntrarCuenta_Err
        

100     If CheckMAC(MacAddress) Then
102         Call WriteShowMessageBox(UserIndex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0001")
            Exit Function

        End If
    
104     If CheckHD(HDserial) Then
106         Call WriteShowMessageBox(UserIndex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0002")
            Exit Function

        End If

108     If Not CheckMailString(CuentaEmail) Then
110         Call WriteShowMessageBox(UserIndex, "Email inválido.")
            Exit Function

        End If
    
112     If Database_Enabled Then
114         EntrarCuenta = EnterAccountDatabase(UserIndex, CuentaEmail, SDesencriptar(CuentaPassword), MacAddress, HDserial, UserList(UserIndex).ip)
    
        Else

116         If CuentaExiste(CuentaEmail) Then
118             If Not ObtenerBaneo(CuentaEmail) Then

                    Dim PasswordHash As String, Salt As String

120                 PasswordHash = GetVar(CuentasPath & UCase$(CuentaEmail) & ".act", "INIT", "PASSWORD")
122                 Salt = GetVar(CuentasPath & UCase$(CuentaEmail) & ".act", "INIT", "SALT")

124                 If PasswordValida(SDesencriptar(CuentaPassword), PasswordHash, Salt) Then
126                     If ObtenerValidacion(CuentaEmail) Then
128                         Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "MacAdress", MacAddress)
130                         Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "HDserial", HDserial)
132                         Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "UltimoAcceso", Date & " " & Time)
134                         Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "UltimaIP", UserList(UserIndex).ip)
                        
136                         UserList(UserIndex).Cuenta = CuentaEmail
                        
138                         EntrarCuenta = True
                        Else
140                         Call WriteShowMessageBox(UserIndex, "¡La cuenta no ha sido validada aún!")

                        End If

                    Else
142                     Call WriteShowMessageBox(UserIndex, "Contraseña inválida.")

                    End If

                Else
144                 Call WriteShowMessageBox(UserIndex, "La cuenta se encuentra baneada debido a: " & ObtenerMotivoBaneo(CuentaEmail) & ". Esta decisión fue tomada por: " & ObtenerQuienBaneo(CuentaEmail) & ".")

                End If

            Else
146             Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")

            End If

        End If
    
        
        Exit Function

EntrarCuenta_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EntrarCuenta", Erl)
        Resume Next
        
End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef name As String, ByRef UserCuenta As String)

    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim n    As Integer

        Dim tStr As String
        
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .name & " ha intentado loguear a " & name & " desde la IP " & .ip)
            
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            
            Exit Sub

        End If
        
        If MaxUsersPorCuenta > 0 Then
            If GetUsersLoggedAccountDatabase(.AccountID) >= MaxUsersPorCuenta Then
                If MaxUsersPorCuenta = 1 Then
                    Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                Else
                    Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")

                End If

                
                Exit Sub

            End If

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
            
            'Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
        '¿Este IP ya esta conectado?
        If MaxConexionesIP > 0 Then
            If ContarMismaIP(UserIndex, .ip) >= MaxConexionesIP Then
                Call WriteShowMessageBox(UserIndex, "Has alcanzado el límite de conexiones por IP.")
                Exit Sub
            End If
        End If
        
        '¿Supera el máximo de usuarios por cuenta?
        If MaxUsersPorCuenta > 0 Then
            If QueryData!logged >= MaxUsersPorCuenta Then
                If MaxUsersPorCuenta = 1 Then
                    Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                Else
                    Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")

                End If

                Exit Sub

            End If

        End If

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
        '
        'Call CloseSocket(UserIndex)
        'Exit Sub
        '    End If
        'End If
        
        If ServerSoloGMs > 0 Then
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
                ' Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                
                'Call CloseSocket(UserIndex)
                Exit Sub

            End If

        End If
        
        '¿Ya esta conectado el personaje?
        If CheckForSameName(name) Then
            If UserList(NameIndex(name)).Counters.Saliendo Then
                Call WriteShowMessageBox(UserIndex, "El usuario está saliendo.")
            Else
                Call WriteShowMessageBox(UserIndex, "Perdon, un usuario con el mismo nombre se ha logueado.")

            End If

            
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
        
        Call LoadUserIntervals(UserIndex)
        Call WriteIntervals(UserIndex)
        
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
        
        'Mapa válido
        If Not MapaValido(.Pos.Map) Then
            Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
            
            Call CloseSocket(UserIndex)
            Exit Sub

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
        
        If Not (.flags.Privilegios And PlayerType.user) Then
            Call DoAdminInvisible(UserIndex)
        End If
        
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
        
        If NumUsers > RecordUsuarios Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultáneamente: " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
            RecordUsuarios = NumUsers
            
            If Database_Enabled Then
                Call SetRecordUsersDatabase(RecordUsuarios)
            Else
                Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(RecordUsuarios))
            End If
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
    
    
    'N = FreeFile
    'Log
    'Open App.Path & "\logs\Connect.log" For Append Shared As #N
    'Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
    'Close #N

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
        
        On Error GoTo SendMOTD_Err
        

        Dim j As Long

100     For j = 1 To MaxLines
102         Call WriteConsoleMsg(UserIndex, MOTD(j).texto, FontTypeNames.FONTTYPE_EXP)
104     Next j
    
        
        Exit Sub

SendMOTD_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.SendMOTD", Erl)
        Resume Next
        
End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
        
        On Error GoTo ResetFacciones_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
100     With UserList(UserIndex).Faccion
102         .ArmadaReal = 0
104         .CiudadanosMatados = 0
106         .CriminalesMatados = 0
108         .Status = 0
110         .FuerzasCaos = 0
112         .FechaIngreso = ""
114         .RecibioArmaduraCaos = 0
116         .RecibioArmaduraReal = 0
118         .RecibioExpInicialCaos = 0
120         .RecibioExpInicialReal = 0
122         .RecompensasCaos = 0
124         .RecompensasReal = 0
126         .Reenlistadas = 0
128         .NivelIngreso = 0
130         .MatadosIngreso = 0
132         .NextRecompensa = 0

        End With

        
        Exit Sub

ResetFacciones_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetFacciones", Erl)
        Resume Next
        
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
        
        On Error GoTo ResetContadores_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '05/20/2007 Integer - Agregue todas las variables que faltaban.
        '*************************************************
100     With UserList(UserIndex).Counters
102         .AGUACounter = 0
104         .AttackCounter = 0
106         .Ceguera = 0
108         .COMCounter = 0
110         .Estupidez = 0
112         .Frio = 0
114         .HPCounter = 0
116         .IdleCount = 0
118         .Invisibilidad = 0
120         .Paralisis = 0
122         .Inmovilizado = 0
124         .Pasos = 0
126         .Pena = 0
128         .PiqueteC = 0
130         .STACounter = 0
132         .Veneno = 0
134         .Trabajando = 0
136         .Ocultando = 0
138         .Lava = 0
140         .Maldicion = 0
142         .Saliendo = False
144         .Salir = 0
146         .TiempoOculto = 0
148         .TimerMagiaGolpe = 0
150         .TimerGolpeMagia = 0
152         .TimerLanzarSpell = 0
154         .TimerPuedeAtacar = 0
156         .TimerPuedeUsarArco = 0
158         .TimerPuedeTrabajar = 0
160         .TimerUsar = 0
            'Ladder
162         .Incineracion = 0
            'Ladder
164         .ScrollExperiencia = 0
166         .ScrollOro = 0
168         .Oxigeno = 0
170         .TiempoParaSubastar = 0
172         .TimerPerteneceNpc = 0
174         .TimerPuedeSerAtacado = 0

        End With

        
        Exit Sub

ResetContadores_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetContadores", Erl)
        Resume Next
        
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        
        On Error GoTo ResetCharInfo_Err
        

100     With UserList(UserIndex).Char
102         .Body = 0
104         .CascoAnim = 0
106         .CharIndex = 0
108         .FX = 0
110         .Head = 0
112         .loops = 0
114         .heading = 0
116         .loops = 0
118         .ShieldAnim = 0
120         .WeaponAnim = 0
122         .Arma_Aura = ""
124         .Body_Aura = ""
126         .Head_Aura = ""
128         .Otra_Aura = ""
130         .Escudo_Aura = ""
132         .ParticulaFx = 0
134         .speeding = VelocidadCero

        End With

        
        Exit Sub

ResetCharInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetCharInfo", Erl)
        Resume Next
        
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
        
        On Error GoTo ResetBasicUserInfo_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        'Agregue que se resetee el maná
        '*************************************************
        Dim LoopC As Integer

100     With UserList(UserIndex)
102         .name = vbNullString
104         .Cuenta = vbNullString
106         .Id = -1
108         .AccountID = -1
110         .modName = vbNullString
112         .Desc = vbNullString
114         .DescRM = vbNullString
116         .Pos.Map = 0
118         .Pos.x = 0
120         .Pos.Y = 0
122         .ip = vbNullString
124         .clase = 0
126         .email = vbNullString
128         .genero = 0
130         .Hogar = 0
132         .raza = 0
134         .EmpoCont = 0
        
            'Ladder     Reseteo de Correos
136         .Correo.CantCorreo = 0
138         .Correo.NoLeidos = 0
        
140         For LoopC = 1 To MAX_CORREOS_SLOTS
142             .Correo.Mensaje(LoopC).Remitente = ""
144             .Correo.Mensaje(LoopC).Mensaje = ""
146             .Correo.Mensaje(LoopC).Item = 0
148             .Correo.Mensaje(LoopC).ItemCount = 0
150             .Correo.Mensaje(LoopC).Fecha = ""
152             .Correo.Mensaje(LoopC).Leido = 0
154         Next LoopC

            'Ladder     Reseteo de Correos
        
156         With .Stats
158             .InventLevel = 0
160             .Banco = 0
162             .ELV = 0
164             .ELU = 0
166             .Exp = 0
168             .def = 0
                '.CriminalesMatados = 0
170             .NPCsMuertos = 0
172             .UsuariosMatados = 0
174             .SkillPts = 0
176             .GLD = 0
178             .UserAtributos(1) = 0
180             .UserAtributos(2) = 0
182             .UserAtributos(3) = 0
184             .UserAtributos(4) = 0
186             .UserAtributosBackUP(1) = 0
188             .UserAtributosBackUP(2) = 0
190             .UserAtributosBackUP(3) = 0
192             .UserAtributosBackUP(4) = 0
194             .MaxMAN = 0
196             .MinMAN = 0
            
            End With
            
            .NroMascotas = 0
        
        End With

        
        Exit Sub

ResetBasicUserInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetBasicUserInfo", Erl)
        Resume Next
        
End Sub

Sub ResetGuildInfo(ByVal UserIndex As Integer)
        
        On Error GoTo ResetGuildInfo_Err
        

100     If UserList(UserIndex).EscucheClan > 0 Then
102         Call modGuilds.GMDejaDeEscucharClan(UserIndex, UserList(UserIndex).EscucheClan)
104         UserList(UserIndex).EscucheClan = 0

        End If

106     If UserList(UserIndex).GuildIndex > 0 Then
108         Call modGuilds.m_DesconectarMiembroDelClan(UserIndex, UserList(UserIndex).GuildIndex)

        End If

110     UserList(UserIndex).GuildIndex = 0
    
        
        Exit Sub

ResetGuildInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetGuildInfo", Erl)
        Resume Next
        
End Sub

Sub ResetUserFlags(ByVal UserIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/29/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '03/29/2006 Maraxus - Reseteo el CentinelaOK también.
        '*************************************************
        
        On Error GoTo ResetUserFlags_Err
        

100     With UserList(UserIndex).flags
102         .LevelBackup = 0
104         .Comerciando = False
106         .Ban = 0
108         .Escondido = 0
110         .DuracionEfecto = 0
112         .ScrollExp = 1
114         .ScrollOro = 1
116         .NpcInv = 0
118         .StatsChanged = 0
120         .TargetNPC = 0
122         .TargetNpcTipo = eNPCType.Comun
124         .TargetObj = 0
126         .TargetObjMap = 0
128         .TargetObjX = 0
130         .TargetObjY = 0
132         .TargetUser = 0
134         .TipoPocion = 0
136         .TomoPocion = False
138         .Descuento = vbNullString
140         .Hambre = 0
142         .Sed = 0
144         .Descansar = False
146         .Navegando = 0
148         .Oculto = 0
150         .Envenenado = 0
152         .Ahogandose = 0
154         .invisible = 0
156         .Paralizado = 0
158         .Inmovilizado = 0
160         .Maldicion = 0
162         .Bendicion = 0
164         .Meditando = 0
166         .Escribiendo = 0
168         .Privilegios = 0
170         .PuedeMoverse = 0
172         .OldBody = 0
174         .OldHead = 0
176         .AdminInvisible = 0
178         .ValCoDe = 0
180         .Hechizo = 0
182         .TimesWalk = 0
184         .StartWalk = 0
186         .CountSH = 0
188         .Silenciado = 0
190         .CentinelaOK = False
192         .AdminPerseguible = False
            'Ladder
194         .VecesQueMoriste = 0
196         .MinutosRestantes = 0
198         .SegundosPasados = 0
200         .RetoA = 0
202         .SolicitudPendienteDe = 0
204         .CarroMineria = 0
206         .DañoMagico = 0
208         .Montado = 0
210         .Incinerado = 0
212         .Casado = 0
214         .Pareja = ""
216         .Candidato = 0
218         .UsandoMacro = False
220         .pregunta = 0
            'Ladder
222         .BattleModo = 0
224         .ResistenciaMagica = 0

226         .Subastando = False
228         .Paraliza = 0
230         .Envenena = 0
232         .NoPalabrasMagicas = 0
234         .NoMagiaEfeceto = 0
236         .incinera = 0
238         .Estupidiza = 0
240         .GolpeCertero = 0
242         .PendienteDelExperto = 0
244         .CarroMineria = 0
246         .PendienteDelSacrificio = 0
248         .AnilloOcultismo = 0
250         .RegeneracionMana = 0
252         .RegeneracionHP = 0
254         .RegeneracionSta = 0
256         .NecesitaOxigeno = False
258         .LastCrimMatado = ""
260         .LastCiudMatado = ""
        
262         .UserLogged = False
264         .FirstPacket = False
        
        End With

        
        Exit Sub

ResetUserFlags_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserFlags", Erl)
        Resume Next
        
End Sub

Sub ResetAccionesPendientes(ByVal UserIndex As Integer)
        
        On Error GoTo ResetAccionesPendientes_Err
        

        '*************************************************
        '*************************************************
100     With UserList(UserIndex).accion
102         .AccionPendiente = False
104         .HechizoPendiente = 0
106         .RunaObj = 0
108         .Particula = 0
110         .TipoAccion = 0
112         .ObjSlot = 0

        End With

        
        Exit Sub

ResetAccionesPendientes_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetAccionesPendientes", Erl)
        Resume Next
        
End Sub

Sub ResetDonadorFlag(ByVal UserIndex As Integer)
        
        On Error GoTo ResetDonadorFlag_Err
        

        '*************************************************
        '*************************************************
100     With UserList(UserIndex).donador
102         .activo = 0
104         .CreditoDonador = 0
106         .FechaExpiracion = 0

        End With

        
        Exit Sub

ResetDonadorFlag_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetDonadorFlag", Erl)
        Resume Next
        
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSpells_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To MAXUSERHECHIZOS
102         UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
            ' UserList(UserIndex).Stats.UserHechizosInterval(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSpells_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSpells", Erl)
        Resume Next
        
End Sub

Sub ResetUserSkills(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSkills_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMSKILLS
102         UserList(UserIndex).Stats.UserSkills(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSkills_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSkills", Erl)
        Resume Next
        
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserBanco_Err
        

        Dim LoopC As Long
    
100     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
102         UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
104         UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
106         UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = 0
108     Next LoopC
    
110     UserList(UserIndex).BancoInvent.NroItems = 0

        
        Exit Sub

ResetUserBanco_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserBanco", Erl)
        Resume Next
        
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
        
        On Error GoTo LimpiarComercioSeguro_Err
        

100     With UserList(UserIndex).ComUsu

102         If .DestUsu > 0 Then
104             Call FinComerciarUsu(.DestUsu)
106             Call FinComerciarUsu(UserIndex)

            End If

        End With

        
        Exit Sub

LimpiarComercioSeguro_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.LimpiarComercioSeguro", Erl)
        Resume Next
        
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSlot_Err
        

100     UserList(UserIndex).ConnIDValida = False
102     UserList(UserIndex).ConnID = -1

104     If UserList(UserIndex).Grupo.Lider = UserIndex Then
106         Call FinalizarGrupo(UserIndex)

        End If

108     If UserList(UserIndex).Grupo.EnGrupo Then
110         Call SalirDeGrupoForzado(UserIndex)

        End If

112     UserList(UserIndex).Grupo.CantidadMiembros = 0
114     UserList(UserIndex).Grupo.EnGrupo = False
116     UserList(UserIndex).Grupo.Lider = 0
118     UserList(UserIndex).Grupo.PropuestaDe = 0
120     UserList(UserIndex).Grupo.Miembros(6) = 0
122     UserList(UserIndex).Grupo.Miembros(1) = 0
124     UserList(UserIndex).Grupo.Miembros(2) = 0
126     UserList(UserIndex).Grupo.Miembros(3) = 0
128     UserList(UserIndex).Grupo.Miembros(4) = 0
130     UserList(UserIndex).Grupo.Miembros(5) = 0

132     Call ResetQuestStats(UserIndex)
134     Call ResetGuildInfo(UserIndex)
136     Call LimpiarComercioSeguro(UserIndex)
138     Call ResetFacciones(UserIndex)
140     Call ResetContadores(UserIndex)
142     Call ResetCharInfo(UserIndex)
144     Call ResetBasicUserInfo(UserIndex)
146     Call ResetUserFlags(UserIndex)
148     Call ResetAccionesPendientes(UserIndex)
150     Call ResetDonadorFlag(UserIndex)
152     Call LimpiarInventario(UserIndex)
154     Call ResetUserSpells(UserIndex)
        'Call ResetUserPets(UserIndex)
156     Call ResetUserBanco(UserIndex)
158     Call ResetUserSkills(UserIndex)

160     With UserList(UserIndex).ComUsu
162         .Acepto = False
164         .cant = 0
166         .DestNick = vbNullString
168         .DestUsu = 0
170         .Objeto = 0

        End With

        
        Exit Sub

ResetUserSlot_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSlot", Erl)
        Resume Next
        
End Sub

Sub CloseUser(ByVal UserIndex As Integer)

    'Call LogTarea("CloseUser " & UserIndex)
    On Error GoTo Errhandler
    
    Dim Map As Integer

    Dim aN  As Integer
    
    Map = UserList(UserIndex).Pos.Map
    
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
    
    errordesc = "ERROR AL DESMONTAR"

    If UserList(UserIndex).flags.Montado > 0 Then
        Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)

    End If
    
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
    Call EraseUserChar(UserIndex, True)
    
    errordesc = "ERROR Update Map Users"
    'Update Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    
    If MapInfo(Map).NumUsers < 0 Then
        MapInfo(Map).NumUsers = 0

    End If
    
    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    'If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)
    errordesc = "ERROR AL RESETSLOT Name:" & UserList(UserIndex).name & " cuenta:" & UserList(UserIndex).Cuenta
    Call ResetUserSlot(UserIndex)
    
    Exit Sub
    
Errhandler:
    Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.description & ". Detalle:" & errordesc)

    Resume Next ' TODO: Provisional hasta solucionar bugs graves

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
        
        On Error GoTo EcharPjsNoPrivilegiados_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To LastUser

102         If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
104             If UserList(LoopC).flags.Privilegios And PlayerType.user Then
106                 Call CloseSocket(LoopC)

                End If

            End If

108     Next LoopC

        
        Exit Sub

EcharPjsNoPrivilegiados_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EcharPjsNoPrivilegiados", Erl)
        Resume Next
        
End Sub
