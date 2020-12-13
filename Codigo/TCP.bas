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

Sub DarCuerpo(ByVal Userindex As Integer)
        
        On Error GoTo DarCuerpo_Err
        

        '*************************************************
        'Author: Nacho (Integer)
        'Last modified: 14/03/2007
        'Elije una cabeza para el usuario y le da un body
        '*************************************************
        Dim NewBody    As Integer

        Dim UserRaza   As Byte

        Dim UserGenero As Byte

100     UserGenero = UserList(Userindex).genero
102     UserRaza = UserList(Userindex).raza

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

156     UserList(Userindex).Char.Body = NewBody

        
        Exit Sub

DarCuerpo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.DarCuerpo", Erl)
        Resume Next
        
End Sub

Sub AsignarAtributos(ByVal Userindex As String)
        
        On Error GoTo AsignarAtributos_Err
        

100     Select Case UserList(Userindex).raza

            Case eRaza.Humano
102             UserList(Userindex).Stats.UserAtributos(1) = 19
104             UserList(Userindex).Stats.UserAtributos(2) = 19
106             UserList(Userindex).Stats.UserAtributos(3) = 19
108             UserList(Userindex).Stats.UserAtributos(4) = 20

110         Case eRaza.Elfo
112             UserList(Userindex).Stats.UserAtributos(1) = 18
114             UserList(Userindex).Stats.UserAtributos(2) = 20
116             UserList(Userindex).Stats.UserAtributos(3) = 21
118             UserList(Userindex).Stats.UserAtributos(4) = 18

120         Case eRaza.Drow
122             UserList(Userindex).Stats.UserAtributos(1) = 20
124             UserList(Userindex).Stats.UserAtributos(2) = 18
126             UserList(Userindex).Stats.UserAtributos(3) = 20
128             UserList(Userindex).Stats.UserAtributos(4) = 19

130         Case eRaza.Gnomo
132             UserList(Userindex).Stats.UserAtributos(1) = 13
134             UserList(Userindex).Stats.UserAtributos(2) = 21
136             UserList(Userindex).Stats.UserAtributos(3) = 22
138             UserList(Userindex).Stats.UserAtributos(4) = 17

140         Case eRaza.Enano
142             UserList(Userindex).Stats.UserAtributos(1) = 21
144             UserList(Userindex).Stats.UserAtributos(2) = 17
146             UserList(Userindex).Stats.UserAtributos(3) = 12
148             UserList(Userindex).Stats.UserAtributos(4) = 22

150         Case eRaza.Orco
152             UserList(Userindex).Stats.UserAtributos(1) = 23
154             UserList(Userindex).Stats.UserAtributos(2) = 17
156             UserList(Userindex).Stats.UserAtributos(3) = 12
158             UserList(Userindex).Stats.UserAtributos(4) = 21

        End Select

        
        Exit Sub

AsignarAtributos_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.AsignarAtributos", Erl)
        Resume Next
        
End Sub

Sub RellenarInventario(ByVal Userindex As String)
        
        On Error GoTo RellenarInventario_Err
        

100     With UserList(Userindex)
        
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
158         .Invent.Object(NumItems).ObjIndex = 2332 ' Pocion violeta
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

Function DescripcionValida(ByVal cad As String) As Boolean
        
        On Error GoTo AsciiValidos_Err
        

        Dim car As Byte

        Dim i   As Integer

100     cad = LCase$(cad)

102     For i = 1 To Len(cad)
104         car = Asc(mid$(cad, i, 1))
    
106         If car < 32 Or car >= 126 Then
108             DescripcionValida = False
                Exit Function

            End If
    
110     Next i

112     DescripcionValida = True

        
        Exit Function

AsciiValidos_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.DescripcionValida", Erl)
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

Function ValidateSkills(ByVal Userindex As Integer) As Boolean
        
        On Error GoTo ValidateSkills_Err
        

        Dim LoopC As Integer

100     For LoopC = 1 To NUMSKILLS

102         If UserList(Userindex).Stats.UserSkills(LoopC) < 0 Then
                Exit Function

104             If UserList(Userindex).Stats.UserSkills(LoopC) > 100 Then UserList(Userindex).Stats.UserSkills(LoopC) = 100

            End If

106     Next LoopC

108     ValidateSkills = True
    
        
        Exit Function

ValidateSkills_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ValidateSkills", Erl)
        Resume Next
        
End Function

Sub ConnectNewUser(ByVal Userindex As Integer, ByRef name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Head As Integer, ByRef UserCuenta As String, ByVal Hogar As eCiudad)
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
102         Call WriteErrorMsg(Userindex, "Nombre invalido.")
            Exit Sub

        End If
    
104     If UserList(Userindex).flags.UserLogged Then
106         Call LogCheating("El usuario " & UserList(Userindex).name & " ha intentado crear a " & name & " desde la IP " & UserList(Userindex).ip)
108         Call CloseSocketSL(Userindex)
110         Call Cerrar_Usuario(Userindex)
            Exit Sub

        End If
    
        Dim LoopC As Long
    
        '¿Existe el personaje?
112     If PersonajeExiste(name) Then
114         Call WriteErrorMsg(Userindex, "Ya existe el personaje.")
            Exit Sub

        End If
    
        'Prevenimos algun bug con dados inválidos
116     If UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then Exit Sub
    
118     UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
120     UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
122     UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
124     UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
        UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma) = UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
    
126     UserList(Userindex).flags.Muerto = 0
128     UserList(Userindex).flags.Escondido = 0

130     UserList(Userindex).flags.Casado = 0
132     UserList(Userindex).flags.Pareja = ""

134     UserList(Userindex).name = name
136     UserList(Userindex).clase = UserClase
138     UserList(Userindex).raza = UserRaza
    
140     UserList(Userindex).Char.Head = Head
    
142     UserList(Userindex).genero = UserSexo
144     UserList(Userindex).Hogar = Hogar
    
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
146     UserList(Userindex).Stats.SkillPts = 10
    
148     UserList(Userindex).Char.Heading = eHeading.SOUTH
    
150     Call DarCuerpo(Userindex) 'Ladder REVISAR
    
152     UserList(Userindex).OrigChar = UserList(Userindex).Char

154     UserList(Userindex).Char.WeaponAnim = NingunArma
156     UserList(Userindex).Char.ShieldAnim = NingunEscudo
158     UserList(Userindex).Char.CascoAnim = NingunCasco

        'Call AsignarAtributos(UserIndex)

        Dim MiInt As Integer
    
160     MiInt = RandomNumber(1, UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
162     UserList(Userindex).Stats.MaxHp = 15 + MiInt
164     UserList(Userindex).Stats.MinHp = 15 + MiInt
    
166     MiInt = RandomNumber(1, UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)

168     If MiInt = 1 Then MiInt = 2
    
170     UserList(Userindex).Stats.MaxSta = 20 * MiInt
172     UserList(Userindex).Stats.MinSta = 20 * MiInt
    
174     UserList(Userindex).Stats.MaxAGU = 100
176     UserList(Userindex).Stats.MinAGU = 100
    
178     UserList(Userindex).Stats.MaxHam = 100
180     UserList(Userindex).Stats.MinHam = 100

182     UserList(Userindex).flags.ScrollExp = 1
184     UserList(Userindex).flags.ScrollOro = 1
    
        '<-----------------MANA----------------------->
186     If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
188         MiInt = UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
190         UserList(Userindex).Stats.MaxMAN = MiInt
192         UserList(Userindex).Stats.MinMAN = MiInt
194     ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid Or UserClase = eClass.Bard Then
196         UserList(Userindex).Stats.MaxMAN = 50
198         UserList(Userindex).Stats.MinMAN = 50
        End If

200     UserList(Userindex).flags.VecesQueMoriste = 0
202     UserList(Userindex).flags.Montado = 0

204     UserList(Userindex).Stats.MaxHit = 2
206     UserList(Userindex).Stats.MinHIT = 1
    
208     UserList(Userindex).Stats.GLD = 0
    
210     UserList(Userindex).Stats.Exp = 0
212     UserList(Userindex).Stats.ELU = 300
214     UserList(Userindex).Stats.ELV = 1
    
216     Call RellenarInventario(Userindex)

        #If ConUpTime Then
218         UserList(Userindex).LogOnTime = Now
220         UserList(Userindex).UpTime = 0
        #End If
    
        'Valores Default de facciones al Activar nuevo usuario
222     Call ResetFacciones(Userindex)
    
224     UserList(Userindex).Faccion.Status = 1
    
226     UserList(Userindex).ChatCombate = 1
228     UserList(Userindex).ChatGlobal = 1
    
        'Resetamos CORREO
230     UserList(Userindex).Correo.CantCorreo = 0
232     UserList(Userindex).Correo.NoLeidos = 0
        'Resetamos CORREO
    
234     UserList(Userindex).Pos.Map = 37
236     UserList(Userindex).Pos.X = 76
238     UserList(Userindex).Pos.Y = 82
    
240     If Not Database_Enabled Then
242         Call GrabarNuevoPjEnCuentaCharfile(UserCuenta, name)
        End If
    
244     UltimoChar = UCase$(name)
    
246     Call SaveNewUser(Userindex)
248     Call ConnectUser(Userindex, name, UserCuenta)
        
        Exit Sub

ConnectNewUser_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ConnectNewUser", Erl)
        Resume Next
        
End Sub

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal Userindex As Integer)

On Error GoTo Errhandler

    Call FlushBuffer(Userindex)

    If Userindex = LastUser Then

        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1

            If LastUser < 1 Then Exit Do
        Loop

    End If
    
    With UserList(Userindex)
    
        'Call SecurityIp.IpRestarConexion(GetLongIp(.ip))

        If .ConnID <> -1 Then Call CloseSocketSL(Userindex)
    
        'Es el mismo user al que está revisando el centinela??
        'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
        ' y lo podemos loguear
        If Centinela.RevisandoUserIndex = Userindex Then Call modCentinela.CentinelaUserLogout
    
        'mato los comercios seguros
        If .ComUsu.DestUsu > 0 Then
        
            If UserList(.ComUsu.DestUsu).flags.UserLogged Then
            
                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
                
                    Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(.ComUsu.DestUsu)
                    
                End If
    
            End If
    
        End If
    
        'Empty buffer for reuse
        Call .incomingData.ReadASCIIStringFixed(.incomingData.Length)
    
        If .flags.UserLogged Then
            Call CloseUser(Userindex)
        
            If NumUsers > 0 Then NumUsers = NumUsers - 1
            Call MostrarNumUsers
        
        Else
            Call ResetUserSlot(Userindex)
    
        End If
    
        .ConnID = -1
        .ConnIDValida = False
        .NumeroPaquetesPorMiliSec = 0
    
    End With
    

    Exit Sub

Errhandler:
    UserList(Userindex).ConnID = -1
    UserList(Userindex).ConnIDValida = False
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0

    Call ResetUserSlot(Userindex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & Userindex)
    Resume Next

End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal Userindex As Integer)
On Error GoTo Errhandler
    
    
    
    UserList(Userindex).ConnID = -1
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0

    If Userindex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(Userindex).flags.UserLogged Then
        If NumUsers <> 0 Then NumUsers = NumUsers - 1
        Call MostrarNumUsers

        Call CloseUser(Userindex)
    End If

    frmMain.Socket2(Userindex).Cleanup
    Unload frmMain.Socket2(Userindex)
    Call ResetUserSlot(Userindex)

Exit Sub

Errhandler:
    UserList(Userindex).ConnID = -1
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(Userindex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal Userindex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo Errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(Userindex).ConnID
    
    'call logindex(UserIndex, "******> Sub CloseSocket. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(Userindex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0

    If Userindex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(Userindex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call MostrarNumUsers
            NURestados = True
            Call CloseUser(Userindex)
    End If
    
    Call ResetUserSlot(Userindex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a CloseSocket del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

Errhandler:
    UserList(Userindex).NumeroPaquetesPorMiliSec = 0
    Call LogError("CLOSESOCKETERR: " & Err.description & " UI:" & Userindex)
    
    If Not NURestados Then
        If UserList(Userindex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
                Call MostrarNumUsers
            End If
            Call LogError("Cerre sin grabar a: " & UserList(Userindex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(Userindex)

End Sub

#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal Userindex As Integer)
        
        On Error GoTo CloseSocketSL_Err
        

        #If UsarQueSocket = 1 Then

100         If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
102             Call BorraSlotSock(UserList(Userindex).ConnID)
104             Call WSApiCloseSocket(UserList(Userindex).ConnID)
106             UserList(Userindex).ConnIDValida = False

            End If

        #ElseIf UsarQueSocket = 0 Then

108         If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
110             frmMain.Socket2(Userindex).Cleanup
112             Unload frmMain.Socket2(Userindex)
114             UserList(Userindex).ConnIDValida = False

            End If

        #ElseIf UsarQueSocket = 2 Then

116         If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
118             Call frmMain.Serv.CerrarSocket(UserList(Userindex).ConnID)
120             UserList(Userindex).ConnIDValida = False

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

Public Sub EnviarDatosASlot(ByVal Userindex As Integer, ByRef Datos As String)
        '***************************************************
        'Author: Unknown
        'Last Modification: 09/11/20
        'Last Modified By: Jopi
        'Se agrega el paquete a la cola, para prevenir errores.
        '***************************************************
        
        On Error GoTo EnviarDatosASlot_Err
        

100     Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(Datos)

        Exit Sub

ErrorHandler:
102     Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & Userindex & "/" & UserList(Userindex).ConnID & "/" & Datos)

        
        Exit Sub

EnviarDatosASlot_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EnviarDatosASlot", Erl)
        Resume Next
        
End Sub

Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean
        
        On Error GoTo EstaPCarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
102         For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

104             If MapData(UserList(Index).Pos.Map, X, Y).Userindex = Index2 Then
106                 EstaPCarea = True
                    Exit Function

                End If
        
108         Next X
110     Next Y

112     EstaPCarea = False

        
        Exit Function

EstaPCarea_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EstaPCarea", Erl)
        Resume Next
        
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
        
        On Error GoTo HayPCarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If X > 0 And Y > 0 And X < 101 And Y < 101 Then
106                 If MapData(Pos.Map, X, Y).Userindex > 0 Then
108                     HayPCarea = True
                        Exit Function

                    End If

                End If

110         Next X
112     Next Y

114     HayPCarea = False

        
        Exit Function

HayPCarea_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.HayPCarea", Erl)
        Resume Next
        
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
        
        On Error GoTo HayOBJarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
106                 HayOBJarea = True
                    Exit Function

                End If
        
108         Next X
110     Next Y

112     HayOBJarea = False

        
        Exit Function

HayOBJarea_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.HayOBJarea", Erl)
        Resume Next
        
End Function

Function ValidateChr(ByVal Userindex As Integer) As Boolean
        
        On Error GoTo ValidateChr_Err
        

100     ValidateChr = UserList(Userindex).Char.Body <> 0 And ValidateSkills(Userindex)

        
        Exit Function

ValidateChr_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ValidateChr", Erl)
        Resume Next
        
End Function

Function EntrarCuenta(ByVal Userindex As Integer, CuentaEmail As String, CuentaPassword As String, MacAddress As String, ByVal HDserial As Long) As Boolean
        
        On Error GoTo EntrarCuenta_Err
        

100     If CheckMAC(MacAddress) Then
102         Call WriteShowMessageBox(Userindex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0001")
            Exit Function

        End If
    
104     If CheckHD(HDserial) Then
106         Call WriteShowMessageBox(Userindex, "Su cuenta se encuentra bajo tolerancia 0. Tiene prohibido el acceso. Cod: #0002")
            Exit Function

        End If

108     If Not CheckMailString(CuentaEmail) Then
110         Call WriteShowMessageBox(Userindex, "Email inválido.")
            Exit Function

        End If
    
112     If Database_Enabled Then
114         EntrarCuenta = EnterAccountDatabase(Userindex, CuentaEmail, SDesencriptar(CuentaPassword), MacAddress, HDserial, UserList(Userindex).ip)
    
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
134                         Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "UltimaIP", UserList(Userindex).ip)
                        
136                         UserList(Userindex).Cuenta = CuentaEmail
                        
138                         EntrarCuenta = True
                        Else
140                         Call WriteShowMessageBox(Userindex, "¡La cuenta no ha sido validada aún!")

                        End If

                    Else
142                     Call WriteShowMessageBox(Userindex, "Contraseña inválida.")

                    End If

                Else
144                 Call WriteShowMessageBox(Userindex, "La cuenta se encuentra baneada debido a: " & ObtenerMotivoBaneo(CuentaEmail) & ". Esta decisión fue tomada por: " & ObtenerQuienBaneo(CuentaEmail) & ".")

                End If

            Else
146             Call WriteShowMessageBox(Userindex, "La cuenta no existe.")

            End If

        End If
    
        
        Exit Function

EntrarCuenta_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.EntrarCuenta", Erl)
        Resume Next
        
End Function

Sub ConnectUser(ByVal Userindex As Integer, ByRef name As String, ByRef UserCuenta As String)

    On Error GoTo Errhandler

    With UserList(Userindex)

        Dim n    As Integer

        Dim tStr As String
        
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .name & " ha intentado loguear a " & name & " desde la IP " & .ip)
            
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(Userindex)
            Call Cerrar_Usuario(Userindex)
            
            Exit Sub
        End If
        
        '¿Supera el máximo de usuarios por cuenta?
        If MaxUsersPorCuenta > 0 Then
            If GetUsersLoggedAccountDatabase(.AccountID) >= MaxUsersPorCuenta Then
                If MaxUsersPorCuenta = 1 Then
                    Call WriteShowMessageBox(Userindex, "Ya hay un usuario conectado con esta cuenta.")
                Else
                    Call WriteShowMessageBox(Userindex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")
                End If

                Call CloseSocket(Userindex)
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
            Call WriteShowMessageBox(Userindex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        '¿Este IP ya esta conectado?
        If MaxConexionesIP > 0 Then
            If ContarMismaIP(Userindex, .ip) >= MaxConexionesIP Then
                Call WriteShowMessageBox(Userindex, "Has alcanzado el límite de conexiones por IP.")
                Call CloseSocket(Userindex)
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
                Call WriteShowMessageBox(Userindex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                
                Call CloseSocket(Userindex)
                Exit Sub
            End If

        End If
        
        '¿Ya esta conectado el personaje?
        If CheckForSameName(name) Then
            If UserList(NameIndex(name)).Counters.Saliendo Then
                Call WriteShowMessageBox(Userindex, "El usuario está saliendo.")
            Else
                Call WriteShowMessageBox(Userindex, "Perdon, un usuario con el mismo nombre se ha logueado.")
            End If
            
            Call CloseSocket(Userindex)
            Exit Sub

        End If
        
        If EnPausa Then
            Call WritePauseToggle(Userindex)
            Call WriteConsoleMsg(Userindex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
            Call CloseSocket(Userindex)
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
        Call LoadUser(Userindex)

        If Not ValidateChr(Userindex) Then
            Call WriteShowMessageBox(Userindex, "Error en el personaje. Comuniquese con el staff.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
    
        If UCase$(.Cuenta) <> UCase$(UserCuenta) Then
            Call WriteShowMessageBox(Userindex, "El personaje no corresponde a su cuenta.")
            Call CloseSocket(Userindex)
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
        
        .CurrentInventorySlots = getMaxInventorySlots(Userindex)
        
        Call WriteInventoryUnlockSlots(Userindex)
        
        Call LoadUserIntervals(Userindex)
        Call WriteIntervals(Userindex)
        
        Call UpdateUserInv(True, Userindex, 0)
        Call UpdateUserHechizos(True, Userindex, 0)
        
        Call EnviarLlaves(Userindex)

        If .Correo.NoLeidos > 0 Then
            Call WriteCorreoPicOn(Userindex)
        End If

        If .flags.Paralizado Then
            Call WriteParalizeOK(Userindex)
        End If
        
        If .flags.Inmovilizado Then
            Call WriteInmovilizaOK(Userindex)
        End If
        
        ''
        'TODO : Feo, esto tiene que ser parche cliente
        If .flags.Estupidez = 0 Then
            Call WriteDumbNoMore(Userindex)
        End If
        
        'Ladder Inmunidad
        .flags.Inmunidad = 1
        .Counters.TiempoDeInmunidad = INTERVALO_INMUNIDAD
        'Ladder Inmunidad
        
        
        
        'Mapa válido
        If Not MapaValido(.Pos.Map) Then
            Call WriteErrorMsg(Userindex, "EL PJ se encuenta en un mapa invalido.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        
        'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
        'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martin Sotuyo Dodero (Maraxus)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

            Dim FoundPlace As Boolean

            Dim esAgua     As Boolean

            Dim tX         As Long

            Dim tY         As Long
        
            FoundPlace = False
            esAgua = (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0
        
            For tY = .Pos.Y - 1 To .Pos.Y + 1
                For tX = .Pos.X - 1 To .Pos.X + 1

                    If esAgua Then

                        'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                        If LegalPos(.Pos.Map, tX, tY, True, False) Then
                            FoundPlace = True
                            Exit For

                        End If

                    Else

                        'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                        If LegalPos(.Pos.Map, tX, tY, False, True) Then
                            FoundPlace = True
                            Exit For

                        End If

                    End If

                Next tX
            
                If FoundPlace Then Exit For
            Next tY
        
            If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
                .Pos.X = tX
                .Pos.Y = tY
            Else

                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex <> 0 Then

                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                    If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu > 0 Then

                        'Le avisamos al que estaba comerciando que se tuvo que ir.
                        If UserList(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu).flags.UserLogged Then
                            Call FinComerciarUsu(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu)
                            Call WriteConsoleMsg(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                        'Lo sacamos.
                        If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex).flags.UserLogged Then
                            Call FinComerciarUsu(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex)
                            Call WriteErrorMsg(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")
                        End If

                    End If
                
                    Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).Userindex)

                End If

            End If

        End If
        
        'If in the water, and has a boat, equip it!
        If .Invent.BarcoObjIndex > 0 And (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Then

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
                Call WriteNadarToggle(Userindex, True)
            Else
                Call WriteNadarToggle(Userindex, False)

            End If
        End If

        Call WriteUserIndexInServer(Userindex) 'Enviamos el User index
        .flags.NecesitaOxigeno = RequiereOxigeno(.Pos.Map)
        
        Call WriteHora(Userindex)
        Call WriteChangeMap(Userindex, .Pos.Map) 'Carga el mapa
        
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
        
        'Crea  el personaje del usuario
        Call MakeUserChar(True, .Pos.Map, Userindex, .Pos.Map, .Pos.X, .Pos.Y, 1)

        Call WriteUserCharIndexInServer(Userindex)
        
        If (.flags.Privilegios And PlayerType.user) = 0 Then
            Call DoAdminInvisible(Userindex)
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        End If
        
        Call WriteVelocidadToggle(Userindex)
        
        Call WriteUpdateUserStats(Userindex)
        
        Call WriteUpdateHungerAndThirst(Userindex)
        
        Call WriteUpdateDM(Userindex)
        Call WriteUpdateRM(Userindex)
        
        Call SendMOTD(Userindex)
        
        Call SetUserLogged(Userindex)
        
        'Actualiza el Num de usuarios
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
        .Counters.LastSave = GetTickCount
        
        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
        If .Stats.SkillPts > 0 Then
            Call WriteSendSkills(Userindex)
            Call WriteLevelUp(Userindex, .Stats.SkillPts)
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
        
        Call WriteFYA(Userindex)
        Call WriteBindKeys(Userindex)
        
        If UserList(Userindex).NroMascotas > 0 And Not MapInfo(UserList(Userindex).Pos.Map).Seguro Then
            Dim i As Integer
            For i = 1 To MAXMASCOTAS
                If UserList(Userindex).MascotasType(i) > 0 Then
                    UserList(Userindex).MascotasIndex(i) = SpawnNpc(UserList(Userindex).MascotasType(i), UserList(Userindex).Pos, True, True)
                    
                    If UserList(Userindex).MascotasIndex(i) > 0 Then
                        Npclist(UserList(Userindex).MascotasIndex(i)).MaestroUser = Userindex
                        Call FollowAmo(UserList(Userindex).MascotasIndex(i))
                    Else
                        UserList(Userindex).MascotasIndex(i) = 0
                    End If
                End If
            Next i
        End If
        
        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(Userindex)
        End If
        
        If .flags.Montado = 1 Then
            .Char.speeding = VelocidadMontura
            Call WriteEquiteToggle(Userindex)
            'Debug.Print "Montado:" & .Char.speeding
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        End If
        
        If .flags.Muerto = 1 Then
            .Char.speeding = VelocidadMuerto
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
        End If
        
        If .GuildIndex > 0 Then

            'welcome to the show baby...
            If Not modGuilds.m_ConectarMiembroAClan(Userindex, .GuildIndex) Then
                Call WriteConsoleMsg(Userindex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
            End If

        End If
        
        tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
    
        If LenB(tStr) <> 0 Then
            Call WriteShowMessageBox(Userindex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
        End If
        
        .flags.SolicitudPendienteDe = 0

        If Lloviendo Then
            Call WriteRainToggle(Userindex)
        End If
        
        If ServidorNublado Then
            Call WriteNubesToggle(Userindex)
        End If

        Call WriteLoggedMessage(Userindex)
        
        If .Stats.ELV = 1 Then
            Call WriteConsoleMsg(Userindex, "¡Bienvenido a las tierras de AO20! ¡" & .name & " que tengas buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
        ElseIf .Stats.ELV < 14 Then
            Call WriteConsoleMsg(Userindex, "¡Bienvenido de nuevo " & .name & "! Actualmente estas en el nivel " & .Stats.ELV & " en " & DarNameMapa(.Pos.Map) & ", ¡buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
        End If

        If Status(Userindex) = 2 Or Status(Userindex) = 0 Then
            Call WriteSafeModeOff(Userindex)
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call WriteSafeModeOn(Userindex)
        End If
        
        'Call modGuilds.SendGuildNews(UserIndex)
        
        If .MENSAJEINFORMACION <> vbNullString Then
            Call WriteConsoleMsg(Userindex, .MENSAJEINFORMACION, FontTypeNames.FONTTYPE_CENTINELA)
            .MENSAJEINFORMACION = vbNullString
        End If

        tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
        
        If LenB(tStr) <> 0 Then
            Call WriteShowMessageBox(Userindex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
        End If

        If EventoActivo Then
            Call WriteConsoleMsg(Userindex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)
        End If
        
        Call WriteContadores(Userindex)
        Call WriteOxigeno(Userindex)

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
    Call WriteShowMessageBox(Userindex, "El personaje contiene un error, comuniquese con un miembro del staff.")
    
    
    'N = FreeFile
    'Log
    'Open App.Path & "\logs\Connect.log" For Append Shared As #N
    'Print #N, UserList(UserIndex).name & " ha entrado al juego. UserIndex:" & UserIndex & " " & Time & " " & Date
    'Close #N

End Sub

Sub SendMOTD(ByVal Userindex As Integer)
        
        On Error GoTo SendMOTD_Err
        

        Dim j As Long

100     For j = 1 To MaxLines
102         Call WriteConsoleMsg(Userindex, MOTD(j).texto, FontTypeNames.FONTTYPE_EXP)
104     Next j
    
        
        Exit Sub

SendMOTD_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.SendMOTD", Erl)
        Resume Next
        
End Sub

Sub ResetFacciones(ByVal Userindex As Integer)
        
        On Error GoTo ResetFacciones_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
        '*************************************************
100     With UserList(Userindex).Faccion
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

Sub ResetContadores(ByVal Userindex As Integer)
        
        On Error GoTo ResetContadores_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '05/20/2007 Integer - Agregue todas las variables que faltaban.
        '*************************************************
100     With UserList(Userindex).Counters
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
            .TiempoDeInmunidad = 0
        End With

        
        Exit Sub

ResetContadores_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetContadores", Erl)
        Resume Next
        
End Sub

Sub ResetCharInfo(ByVal Userindex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '*************************************************
        
        On Error GoTo ResetCharInfo_Err
        

100     With UserList(Userindex).Char
102         .Body = 0
104         .CascoAnim = 0
106         .CharIndex = 0
108         .FX = 0
110         .Head = 0
112         .loops = 0
114         .Heading = 0
116         .loops = 0
118         .ShieldAnim = 0
120         .WeaponAnim = 0
122         .Arma_Aura = ""
124         .Body_Aura = ""
126         .Head_Aura = ""
128         .Otra_Aura = ""
            .Anillo_Aura = ""
130         .Escudo_Aura = ""
132         .ParticulaFx = 0
134         .speeding = VelocidadCero

        End With

        
        Exit Sub

ResetCharInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetCharInfo", Erl)
        Resume Next
        
End Sub

Sub ResetBasicUserInfo(ByVal Userindex As Integer)
        
        On Error GoTo ResetBasicUserInfo_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 03/15/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        'Agregue que se resetee el maná
        '*************************************************
        Dim LoopC As Integer

100     With UserList(Userindex)
102         .name = vbNullString
104         .Cuenta = vbNullString
106         .Id = -1
108         .AccountID = -1
110         .modName = vbNullString
112         .Desc = vbNullString
114         .DescRM = vbNullString
116         .Pos.Map = 0
118         .Pos.X = 0
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

Sub ResetGuildInfo(ByVal Userindex As Integer)
        
        On Error GoTo ResetGuildInfo_Err
        

100     If UserList(Userindex).EscucheClan > 0 Then
102         Call modGuilds.GMDejaDeEscucharClan(Userindex, UserList(Userindex).EscucheClan)
104         UserList(Userindex).EscucheClan = 0

        End If

106     If UserList(Userindex).GuildIndex > 0 Then
108         Call modGuilds.m_DesconectarMiembroDelClan(Userindex, UserList(Userindex).GuildIndex)

        End If

110     UserList(Userindex).GuildIndex = 0
    
        
        Exit Sub

ResetGuildInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetGuildInfo", Erl)
        Resume Next
        
End Sub

Sub ResetUserFlags(ByVal Userindex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 03/29/2006
        'Resetea todos los valores generales y las stats
        '03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
        '03/29/2006 Maraxus - Reseteo el CentinelaOK también.
        '*************************************************
        
        On Error GoTo ResetUserFlags_Err
        

100     With UserList(Userindex).flags
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
208         .Montado = 0
210         .Incinerado = 0
212         .Casado = 0
214         .Pareja = ""
216         .Candidato = 0
218         .UsandoMacro = False
220         .pregunta = 0
            'Ladder
222         .BattleModo = 0

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
258         .LastCrimMatado = vbNullString
260         .LastCiudMatado = vbNullString
        
262         .UserLogged = False
264         .FirstPacket = False
            .Inmunidad = 0
            
            .Mimetizado = 0
            .MascotasGuardadas = 0
        End With

        
        Exit Sub

ResetUserFlags_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserFlags", Erl)
        Resume Next
        
End Sub

Sub ResetAccionesPendientes(ByVal Userindex As Integer)
        
        On Error GoTo ResetAccionesPendientes_Err
        

        '*************************************************
        '*************************************************
100     With UserList(Userindex).Accion
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

Sub ResetDonadorFlag(ByVal Userindex As Integer)
        
        On Error GoTo ResetDonadorFlag_Err
        

        '*************************************************
        '*************************************************
100     With UserList(Userindex).donador
102         .activo = 0
104         .CreditoDonador = 0
106         .FechaExpiracion = 0

        End With

        
        Exit Sub

ResetDonadorFlag_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetDonadorFlag", Erl)
        Resume Next
        
End Sub

Sub ResetUserSpells(ByVal Userindex As Integer)
        
        On Error GoTo ResetUserSpells_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To MAXUSERHECHIZOS
102         UserList(Userindex).Stats.UserHechizos(LoopC) = 0
            ' UserList(UserIndex).Stats.UserHechizosInterval(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSpells_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSpells", Erl)
        Resume Next
        
End Sub

Sub ResetUserSkills(ByVal Userindex As Integer)
        
        On Error GoTo ResetUserSkills_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMSKILLS
102         UserList(Userindex).Stats.UserSkills(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSkills_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSkills", Erl)
        Resume Next
        
End Sub

Sub ResetUserBanco(ByVal Userindex As Integer)
        
        On Error GoTo ResetUserBanco_Err
        

        Dim LoopC As Long
    
100     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
102         UserList(Userindex).BancoInvent.Object(LoopC).Amount = 0
104         UserList(Userindex).BancoInvent.Object(LoopC).Equipped = 0
106         UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex = 0
108     Next LoopC
    
110     UserList(Userindex).BancoInvent.NroItems = 0

        
        Exit Sub

ResetUserBanco_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserBanco", Erl)
        Resume Next
        
End Sub

Sub ResetUserKeys(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim i As Integer
        
        For i = 1 To MAXKEYS
            .Keys(i) = 0
        Next
    End With
End Sub

Public Sub LimpiarComercioSeguro(ByVal Userindex As Integer)
        
        On Error GoTo LimpiarComercioSeguro_Err
        

100     With UserList(Userindex).ComUsu

102         If .DestUsu > 0 Then
104             Call FinComerciarUsu(.DestUsu)
106             Call FinComerciarUsu(Userindex)

            End If

        End With

        
        Exit Sub

LimpiarComercioSeguro_Err:
        Call RegistrarError(Err.Number, Err.description, "TCP.LimpiarComercioSeguro", Erl)
        Resume Next
        
End Sub

Sub ResetUserSlot(ByVal Userindex As Integer)
        
        On Error GoTo ResetUserSlot_Err
        

100     UserList(Userindex).ConnIDValida = False
102     UserList(Userindex).ConnID = -1

104     If UserList(Userindex).Grupo.Lider = Userindex Then
106         Call FinalizarGrupo(Userindex)

        End If

108     If UserList(Userindex).Grupo.EnGrupo Then
110         Call SalirDeGrupoForzado(Userindex)

        End If

112     UserList(Userindex).Grupo.CantidadMiembros = 0
114     UserList(Userindex).Grupo.EnGrupo = False
116     UserList(Userindex).Grupo.Lider = 0
118     UserList(Userindex).Grupo.PropuestaDe = 0
120     UserList(Userindex).Grupo.Miembros(6) = 0
122     UserList(Userindex).Grupo.Miembros(1) = 0
124     UserList(Userindex).Grupo.Miembros(2) = 0
126     UserList(Userindex).Grupo.Miembros(3) = 0
128     UserList(Userindex).Grupo.Miembros(4) = 0
130     UserList(Userindex).Grupo.Miembros(5) = 0

132     Call ResetQuestStats(Userindex)
134     Call ResetGuildInfo(Userindex)
136     Call LimpiarComercioSeguro(Userindex)
138     Call ResetFacciones(Userindex)
140     Call ResetContadores(Userindex)
142     Call ResetCharInfo(Userindex)
144     Call ResetBasicUserInfo(Userindex)
146     Call ResetUserFlags(Userindex)
148     Call ResetAccionesPendientes(Userindex)
150     Call ResetDonadorFlag(Userindex)
152     Call LimpiarInventario(Userindex)
154     Call ResetUserSpells(Userindex)
        'Call ResetUserPets(UserIndex)
156     Call ResetUserBanco(Userindex)
158     Call ResetUserSkills(Userindex)
        Call ResetUserKeys(Userindex)

160     With UserList(Userindex).ComUsu
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

Sub CloseUser(ByVal Userindex As Integer)

    On Error GoTo Errhandler
    
    Dim errordesc As String
    Dim Map As Integer
    Dim aN  As Integer
    Dim i   As Integer
    
    Map = UserList(Userindex).Pos.Map
    
    errordesc = "ERROR AL SETEAR NPC"
    
    aN = UserList(Userindex).flags.AtacadoPorNpc

    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = vbNullString

    End If

    aN = UserList(Userindex).flags.NPCAtacado

    If aN > 0 Then
        If Npclist(aN).flags.AttackedFirstBy = UserList(Userindex).name Then
            Npclist(aN).flags.AttackedFirstBy = vbNullString

        End If

    End If

    UserList(Userindex).flags.AtacadoPorNpc = 0
    UserList(Userindex).flags.NPCAtacado = 0
    
    errordesc = "ERROR AL DESMONTAR"

    If UserList(Userindex).flags.Montado > 0 Then
        Call DoMontar(Userindex, ObjData(UserList(Userindex).Invent.MonturaObjIndex), UserList(Userindex).Invent.MonturaSlot)
    End If
    
    errordesc = "ERROR AL SACAR MIMETISMO"
    If UserList(Userindex).flags.Mimetizado = 1 Then
        UserList(Userindex).Char.Body = UserList(Userindex).CharMimetizado.Body
        UserList(Userindex).Char.Head = UserList(Userindex).CharMimetizado.Head
        UserList(Userindex).Char.CascoAnim = UserList(Userindex).CharMimetizado.CascoAnim
        UserList(Userindex).Char.ShieldAnim = UserList(Userindex).CharMimetizado.ShieldAnim
        UserList(Userindex).Char.WeaponAnim = UserList(Userindex).CharMimetizado.WeaponAnim
        UserList(Userindex).Counters.Mimetismo = 0
        UserList(Userindex).flags.Mimetizado = 0
    End If
    
    errordesc = "ERROR AL ENVIAR PARTICULA"
    
    UserList(Userindex).Char.FX = 0
    UserList(Userindex).Char.loops = 0
    UserList(Userindex).Char.ParticulaFx = 0
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, 0, 0, True))
    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, 0, 0))
    
    UserList(Userindex).flags.UserLogged = False
    UserList(Userindex).Counters.Saliendo = False
    
    errordesc = "ERROR AL ENVIAR INVI"
    
    'Le devolvemos el body y head originales
    If UserList(Userindex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(Userindex)
    
    errordesc = "ERROR AL CANCELAR SUBASTA"

    If UserList(Userindex).flags.Subastando = True Then
        Call CancelarSubasta

    End If
    
    errordesc = "ERROR AL BORRAR INDEX DE TORNEO"

    If UserList(Userindex).flags.EnTorneo = True Then
        Call BorrarIndexInTorneo(Userindex)
        UserList(Userindex).flags.EnTorneo = False

    End If
    
    'Save statistics
    'Call Statistics.UserDisconnected(UserIndex)
    
    ' Grabamos el personaje del usuario
    
    errordesc = "ERROR AL GRABAR PJ"
    
    If UserList(Userindex).flags.BattleModo = 0 Then
        Call SaveUser(Userindex, True)
    Else
        'Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)
        Call SaveBattlePoints(Userindex)

    End If

    errordesc = "ERROR AL DESCONTAR USER DE MAPA"

    If MapInfo(Map).NumUsers > 0 Then
        Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageRemoveCharDialog(UserList(Userindex).Char.CharIndex))

    End If

    errordesc = "ERROR AL ERASEUSERCHAR"
    
    'Borrar el personaje
    Call EraseUserChar(Userindex, True)
    
    errordesc = "ERROR AL BORRAR MASCOTAS"
    
    'Borrar mascotas
    For i = 1 To MAXMASCOTAS
        If UserList(Userindex).MascotasIndex(i) > 0 Then
            If Npclist(UserList(Userindex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(Userindex).MascotasIndex(i))
        End If
    Next i
    
    errordesc = "ERROR Update Map Users"
    
    'Update Map Users
    MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    
    If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0

    ' Si el usuario habia dejado un msg en la gm's queue lo borramos
    'If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)
    
    errordesc = "ERROR AL RESETEAR FLAGS Name:" & UserList(Userindex).name & " cuenta:" & UserList(Userindex).Cuenta
    
    'Reseteo los estados del juagador, fuerza el cierre del cliente.
    Call ResetUserFlags(Userindex)
    
    errordesc = "ERROR AL RESETSLOT Name:" & UserList(Userindex).name & " cuenta:" & UserList(Userindex).Cuenta
    
    Call ResetUserSlot(Userindex)
    
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
