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
158     Call RegistrarError(Err.Number, Err.description, "TCP.DarCuerpo", Erl)
160     Resume Next
        
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
160     Call RegistrarError(Err.Number, Err.description, "TCP.AsignarAtributos", Erl)
162     Resume Next
        
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

                Case eClass.Bard, eClass.Cleric, eClass.Paladin, eClass.Assasin, eClass.Bandit
120                 .Invent.Object(NumItems).ObjIndex = 1617 ' Pocion Azul
122                 .Invent.Object(NumItems).Amount = 50
124                 NumItems = NumItems + 1

            End Select

            ' Arma y hechizos
126         Select Case .clase

                Case eClass.Mage, eClass.Cleric, eClass.Druid, eClass.Bard
128                 .Stats.UserHechizos(1) = 1 ' Dardo mágico

            End Select
        
            ' Pociones amarillas y verdes
130         Select Case .clase

                Case eClass.Assasin, eClass.Bard, eClass.Cleric, eClass.Hunter, eClass.Paladin, eClass.Trabajador, eClass.Warrior, eClass.Bandit, eClass.Pirat, eClass.Thief
132                 .Invent.Object(NumItems).ObjIndex = 1618 ' Pocion Amarilla
134                 .Invent.Object(NumItems).Amount = 25
136                 NumItems = NumItems + 1

138                 .Invent.Object(NumItems).ObjIndex = 1619 ' Pocion Verde
140                 .Invent.Object(NumItems).Amount = 25
142                 NumItems = NumItems + 1

            End Select
            
            ' Poción violeta
144         .Invent.Object(NumItems).ObjIndex = 2332 ' Pocion violeta
146         .Invent.Object(NumItems).Amount = 10
148         NumItems = NumItems + 1
        
            ' Equipo el arma
150         .Invent.Object(NumItems).ObjIndex = 460 ' Daga (Newbies)
152         .Invent.Object(NumItems).Amount = 1
154         .Invent.Object(NumItems).Equipped = 1
156         .Invent.WeaponEqpSlot = NumItems
158         .Invent.WeaponEqpObjIndex = .Invent.Object(NumItems).ObjIndex
160         .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
162         NumItems = NumItems + 1
        
            
164         If .genero = eGenero.Hombre Then
166             If .raza = Enano Or .raza = Gnomo Then
168                 .Invent.Object(NumItems).ObjIndex = 466 'Vestimentas de Bajo (Newbies)
                Else
                
170                 .Invent.Object(NumItems).ObjIndex = RandomNumber(463, 465) ' Vestimentas comunes (Newbies)
                End If
            Else
172             If .raza = Enano Or .raza = Gnomo Then
174                 .Invent.Object(NumItems).ObjIndex = 563 'Vestimentas de Baja (Newbies)
                Else
176                 .Invent.Object(NumItems).ObjIndex = RandomNumber(1283, 1285) ' Vestimentas de Mujer (Newbies)
                End If
            End If
                        
178         .Invent.Object(NumItems).Amount = 1
180         .Invent.Object(NumItems).Equipped = 1
182         .Invent.ArmourEqpSlot = NumItems
184         .Invent.ArmourEqpObjIndex = .Invent.Object(NumItems).ObjIndex
186          NumItems = NumItems + 1

            ' Animación según raza

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
210     Call RegistrarError(Err.Number, Err.description, "TCP.RellenarInventario", Erl)
212     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "TCP.AsciiValidos", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "TCP.DescripcionValida", Erl)
116     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "TCP.Numeric", Erl)
116     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "TCP.NombrePermitido", Erl)
112     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.description, "TCP.ValidateSkills", Erl)
112     Resume Next
        
End Function

Function ConnectNewUser(ByVal UserIndex As Integer, ByRef name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Head As Integer, ByRef UserCuenta As String, ByVal Hogar As eCiudad) As Boolean
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
        
        Dim LoopC As Long
    
100     If UserList(UserIndex).flags.UserLogged Then
102         Call LogCheating("El usuario " & UserList(UserIndex).name & " ha intentado crear a " & name & " desde la IP " & UserList(UserIndex).ip)
104         Call CloseSocketSL(UserIndex)
106         Call Cerrar_Usuario(UserIndex)
            Exit Function
        End If
        
        ' Nombre válido
108     If Not ValidarNombre(name) Then Exit Function
        
110     If Not NombrePermitido(name) Then
112         Call WriteShowMessageBox(UserIndex, "El nombre no está permitido.")
            Exit Function
        End If

        '¿Existe el personaje?
114     If PersonajeExiste(name) Then
116         Call WriteShowMessageBox(UserIndex, "Ya existe el personaje.")
            Exit Function
        End If
        
        ' Raza válida
118     If UserRaza <= 0 Or UserRaza > NUMRAZAS Then Exit Function
        
        ' Género válido
120     If UserSexo < Hombre Or UserSexo > Mujer Then Exit Function
        
        ' Ciudad válida
122     If Hogar <= 0 Or Hogar > NUMCIUDADES Then Exit Function
        
        ' Cabeza válida
124     If Not ValidarCabeza(UserRaza, UserSexo, Head) Then Exit Function
        
        'Prevenimos algun bug con dados inválidos
126     If UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = 0 Then Exit Function
    
128     UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
130     UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
132     UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
134     UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
136     UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
    
138     UserList(UserIndex).flags.Muerto = 0
140     UserList(UserIndex).flags.Escondido = 0

142     UserList(UserIndex).flags.Casado = 0
144     UserList(UserIndex).flags.Pareja = ""

146     UserList(UserIndex).name = name
148     UserList(UserIndex).clase = UserClase
150     UserList(UserIndex).raza = UserRaza
    
152     UserList(UserIndex).Char.Head = Head
    
154     UserList(UserIndex).genero = UserSexo
156     UserList(UserIndex).Hogar = Hogar
    
        '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
158     UserList(UserIndex).Stats.SkillPts = 10
    
160     UserList(UserIndex).Char.Heading = eHeading.SOUTH
    
162     Call DarCuerpo(UserIndex) 'Ladder REVISAR
    
164     UserList(UserIndex).OrigChar = UserList(UserIndex).Char

166     UserList(UserIndex).Char.WeaponAnim = NingunArma
168     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
170     UserList(UserIndex).Char.CascoAnim = NingunCasco

        'Call AsignarAtributos(UserIndex)

        Dim MiInt As Integer
    
172     MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
174     UserList(UserIndex).Stats.MaxHp = 15 + MiInt
176     UserList(UserIndex).Stats.MinHp = 15 + MiInt
    
178     MiInt = RandomNumber(1, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)

180     If MiInt = 1 Then MiInt = 2
    
182     UserList(UserIndex).Stats.MaxSta = 20 * MiInt
184     UserList(UserIndex).Stats.MinSta = 20 * MiInt
    
186     UserList(UserIndex).Stats.MaxAGU = 100
188     UserList(UserIndex).Stats.MinAGU = 100
    
190     UserList(UserIndex).Stats.MaxHam = 100
192     UserList(UserIndex).Stats.MinHam = 100

194     UserList(UserIndex).flags.ScrollExp = 1
196     UserList(UserIndex).flags.ScrollOro = 1
    
        '<-----------------MANA----------------------->
198     If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
200         MiInt = UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
202         UserList(UserIndex).Stats.MaxMAN = MiInt
204         UserList(UserIndex).Stats.MinMAN = MiInt
206     ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid Or UserClase = eClass.Bard Then
208         UserList(UserIndex).Stats.MaxMAN = 50
210         UserList(UserIndex).Stats.MinMAN = 50
        End If

212     UserList(UserIndex).flags.VecesQueMoriste = 0
214     UserList(UserIndex).flags.Montado = 0

216     UserList(UserIndex).Stats.MaxHit = 2
218     UserList(UserIndex).Stats.MinHIT = 1
    
220     UserList(UserIndex).Stats.GLD = 0
    
222     UserList(UserIndex).Stats.Exp = 0
224     UserList(UserIndex).Stats.ELU = 300
226     UserList(UserIndex).Stats.ELV = 1
    
228     Call RellenarInventario(UserIndex)

        #If ConUpTime Then
230         UserList(UserIndex).LogOnTime = Now
232         UserList(UserIndex).UpTime = 0
        #End If
    
        'Valores Default de facciones al Activar nuevo usuario
234     Call ResetFacciones(UserIndex)
    
236     UserList(UserIndex).Faccion.Status = 1
    
238     UserList(UserIndex).ChatCombate = 1
240     UserList(UserIndex).ChatGlobal = 1
    
        'Resetamos CORREO
242     UserList(UserIndex).Correo.CantCorreo = 0
244     UserList(UserIndex).Correo.NoLeidos = 0
        'Resetamos CORREO
    
246     UserList(UserIndex).Pos.Map = 37
248     UserList(UserIndex).Pos.X = 76
250     UserList(UserIndex).Pos.Y = 82
    
252     If Not Database_Enabled Then
254         Call GrabarNuevoPjEnCuentaCharfile(UserCuenta, name)
        End If
    
256     UltimoChar = UCase$(name)
    
258     Call SaveNewUser(UserIndex)

260     ConnectNewUser = True

262     Call ConnectUser(UserIndex, name, UserCuenta)
        
        Exit Function

ConnectNewUser_Err:
264     Call RegistrarError(Err.Number, Err.description, "TCP.ConnectNewUser", Erl)
266     Resume Next
        
End Function

#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub CloseSocket(ByVal UserIndex As Integer)

    On Error GoTo ErrHandler

100     Call FlushBuffer(UserIndex)

102     If UserIndex = LastUser Then

104         Do Until UserList(LastUser).flags.UserLogged
106             LastUser = LastUser - 1

108             If LastUser < 1 Then Exit Do
            Loop

        End If
    
110     With UserList(UserIndex)
    
            'Call SecurityIp.IpRestarConexion(GetLongIp(.ip))

112         If .ConnID <> -1 Then Call CloseSocketSL(UserIndex)
    
            'Es el mismo user al que está revisando el centinela??
            'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
            ' y lo podemos loguear
114         If Centinela.RevisandoUserIndex = UserIndex Then Call modCentinela.CentinelaUserLogout
    
            'mato los comercios seguros
116         If .ComUsu.DestUsu > 0 Then
        
118             If UserList(.ComUsu.DestUsu).flags.UserLogged Then
            
120                 If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                
122                     Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
124                     Call FinComerciarUsu(.ComUsu.DestUsu)
                    
                    End If
    
                End If
    
            End If
    
            'Empty buffer for reuse
126         Call .incomingData.ReadASCIIStringFixed(.incomingData.Length)
    
128         If .flags.UserLogged Then
130             Call CloseUser(UserIndex)
        
132             If NumUsers > 0 Then NumUsers = NumUsers - 1
134             Call MostrarNumUsers
        
            Else
136             Call ResetUserSlot(UserIndex)
    
            End If
    
138         .ConnID = -1
140         .ConnIDValida = False
142         .NumeroPaquetesPorMiliSec = 0
    
        End With
    

        Exit Sub

ErrHandler:
144     UserList(UserIndex).ConnID = -1
146     UserList(UserIndex).ConnIDValida = False
148     UserList(UserIndex).NumeroPaquetesPorMiliSec = 0

150     Call ResetUserSlot(UserIndex)

152     Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)
154     Resume Next

End Sub

#ElseIf UsarQueSocket = 0 Then

Sub CloseSocket(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    
    
    
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

ErrHandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(UserIndex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub CloseSocket(ByVal UserIndex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo ErrHandler

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

ErrHandler:
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
122     Call RegistrarError(Err.Number, Err.description, "TCP.CloseSocketSL", Erl)
124     Resume Next
        
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
            'Modified By: Jopi
            'Last Modified by: WyroX - Si no hay espacio, flusheo el buffer e intento de nuevo
            'Se agrega el paquete a la cola, para prevenir errores.
            '***************************************************
        
            On Error GoTo EnviarDatosASlot_Err
        

100         Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(Datos)

            Exit Sub

EnviarDatosASlot_Err:
102         If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
104             Call FlushBuffer(UserIndex)
106             Resume
            End If
        
End Sub

Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
        
        On Error GoTo EstaPCarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
102         For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

104             If MapData(UserList(index).Pos.Map, X, Y).UserIndex = Index2 Then
106                 EstaPCarea = True
                    Exit Function

                End If
        
108         Next X
110     Next Y

112     EstaPCarea = False

        
        Exit Function

EstaPCarea_Err:
114     Call RegistrarError(Err.Number, Err.description, "TCP.EstaPCarea", Erl)
116     Resume Next
        
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
        
        On Error GoTo HayPCarea_Err
        

        Dim X As Integer, Y As Integer

100     For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
102         For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1

104             If X > 0 And Y > 0 And X < 101 And Y < 101 Then
106                 If MapData(Pos.Map, X, Y).UserIndex > 0 Then
108                     HayPCarea = True
                        Exit Function

                    End If

                End If

110         Next X
112     Next Y

114     HayPCarea = False

        
        Exit Function

HayPCarea_Err:
116     Call RegistrarError(Err.Number, Err.description, "TCP.HayPCarea", Erl)
118     Resume Next
        
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
114     Call RegistrarError(Err.Number, Err.description, "TCP.HayOBJarea", Erl)
116     Resume Next
        
End Function

Function ValidateChr(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo ValidateChr_Err
        

100     ValidateChr = UserList(UserIndex).Char.Body <> 0 And ValidateSkills(UserIndex)

        
        Exit Function

ValidateChr_Err:
102     Call RegistrarError(Err.Number, Err.description, "TCP.ValidateChr", Erl)
104     Resume Next
        
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
148     Call RegistrarError(Err.Number, Err.description, "TCP.EntrarCuenta", Erl)
150     Resume Next
        
End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef name As String, ByRef UserCuenta As String)

          On Error GoTo ErrHandler

100       With UserList(UserIndex)

              Dim n    As Integer

              Dim tStr As String
        
102           If .flags.UserLogged Then
104               Call LogCheating("El usuario " & .name & " ha intentado loguear a " & name & " desde la IP " & .ip)
            
                  'Kick player ( and leave character inside :D )!
106               Call CloseSocketSL(UserIndex)
108               Call Cerrar_Usuario(UserIndex)
            
                  Exit Sub
              End If
            
              '¿Ya esta conectado el personaje?
              Dim tIndex As Integer
110           tIndex = NameIndex(name)

112           If tIndex > 0 And tIndex <> UserIndex Then
114               If UserList(tIndex).Counters.Saliendo Then
116                   Call WriteShowMessageBox(UserIndex, "El personaje está saliendo.")
                  Else
118                  Call WriteShowMessageBox(UserIndex, "El personaje ya está conectado. Espere mientras es desconectado.")

                      ' Le avisamos al usuario que está jugando, en caso de que haya uno
120                  Call WriteShowMessageBox(tIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")
122                  Call Cerrar_Usuario(tIndex)
                  End If
            
124              Call CloseSocket(UserIndex)
                  Exit Sub

              End If
        
              '¿Supera el máximo de usuarios por cuenta?
126         If MaxUsersPorCuenta > 0 Then
128             If GetUsersLoggedAccountDatabase(.AccountID) >= MaxUsersPorCuenta Then
130                 If MaxUsersPorCuenta = 1 Then
132                     Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                      Else
134                     Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")
                      End If

136                 Call CloseSocket(UserIndex)
                      Exit Sub
                  End If

              End If
        
              'Reseteamos los FLAGS
138         .flags.Escondido = 0
140         .flags.TargetNPC = 0
142         .flags.TargetNpcTipo = eNPCType.Comun
144         .flags.TargetObj = 0
146         .flags.TargetUser = 0
148         .Char.FX = 0
        
              'Controlamos no pasar el maximo de usuarios
150         If NumUsers >= MaxUsers Then
152             Call WriteShowMessageBox(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
154             Call CloseSocket(UserIndex)
                  Exit Sub
              End If
        
              '¿Este IP ya esta conectado?
156         If MaxConexionesIP > 0 Then
158             If ContarMismaIP(UserIndex, .ip) >= MaxConexionesIP Then
160                 Call WriteShowMessageBox(UserIndex, "Has alcanzado el límite de conexiones por IP.")
162                 Call CloseSocket(UserIndex)
                      Exit Sub
                  End If
              End If

              'Reseteamos los privilegios
164         .flags.Privilegios = 0
        
              'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
166         If EsAdmin(name) Then
168             .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
170             Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
                  '    Call LogGM(name, "Se conecto con ip:" & .ip)
172         ElseIf EsDios(name) Then
174             .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
176             Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
                  '    Call LogGM(name, "Se conecto con ip:" & .ip)
178         ElseIf EsSemiDios(name) Then
180             .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
182             Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
                  '    Call LogGM(name, "Se conecto con ip:" & .ip)
184         ElseIf EsConsejero(name) Then
186             .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
188             Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
                  '    Call LogGM(name, "Se conecto con ip:" & .ip)
              Else
190             .flags.Privilegios = .flags.Privilegios Or PlayerType.user
192             .flags.AdminPerseguible = True
              End If

              'Add RM flag if needed
194         If EsRolesMaster(name) Then
196             .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
198             Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg("Servidor> " & name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
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
        
200         If ServerSoloGMs > 0 Then
202             If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
                      ' Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
204                 Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                
206                 Call CloseSocket(UserIndex)
                      Exit Sub
                  End If

              End If
        
208         If EnPausa Then
210             Call WritePauseToggle(UserIndex)
212             Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
214             Call CloseSocket(UserIndex)
                  Exit Sub
              End If
    
              'Donador
216         If DonadorCheck(UserCuenta) Then

                  Dim LoopC As Integer

218             For LoopC = 1 To Donadores.Count

220                 If UCase$(Donadores(LoopC).name) = UCase$(UserCuenta) Then
222                     .donador.activo = 1
224                     .donador.FechaExpiracion = Donadores(LoopC).FechaExpiracion
                          Exit For

                      End If

226             Next LoopC

              End If
        
              ' Seteamos el nombre
228         .name = name
        
              ' Cargamos el personaje
230         Call LoadUser(UserIndex)

232         If Not ValidateChr(UserIndex) Then
234             Call WriteShowMessageBox(UserIndex, "Error en el personaje. Comuniquese con el staff.")
236             Call CloseSocket(UserIndex)
                  Exit Sub
              End If
    
238         If UCase$(.Cuenta) <> UCase$(UserCuenta) Then
240             Call WriteShowMessageBox(UserIndex, "El personaje no corresponde a su cuenta.")
242             Call CloseSocket(UserIndex)
                  Exit Sub
              End If
        
244         If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
246         If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
248         If .Invent.WeaponEqpSlot = 0 And .Invent.NudilloSlot = 0 And .Invent.HerramientaEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        
              'If (.flags.Muerto = 0) Then
              '    .flags.SeguroResu = False
              '    Call WritePartySafeOff(UserIndex)
              'Else
250         .flags.SeguroParty = True
              '    Call WritePartySafeOn(UserIndex)
              'End If
        
252         .flags.SeguroClan = True
        
254         .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
        
256         Call WriteInventoryUnlockSlots(UserIndex)
        
258         Call LoadUserIntervals(UserIndex)
260         Call WriteIntervals(UserIndex)
        
262         Call UpdateUserInv(True, UserIndex, 0)
264         Call UpdateUserHechizos(True, UserIndex, 0)
        
266         Call EnviarLlaves(UserIndex)

268         If .Correo.NoLeidos > 0 Then
270             Call WriteCorreoPicOn(UserIndex)
              End If

272         If .flags.Paralizado Then
274             Call WriteParalizeOK(UserIndex)
              End If
        
276         If .flags.Inmovilizado Then
278             Call WriteInmovilizaOK(UserIndex)
              End If
        
              ''
              'TODO : Feo, esto tiene que ser parche cliente
280         If .flags.Estupidez = 0 Then
282             Call WriteDumbNoMore(UserIndex)
              End If
        
              'Ladder Inmunidad
284         .flags.Inmunidad = 1
286         .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
              'Ladder Inmunidad
        
        
        
              'Mapa válido
288         If Not MapaValido(.Pos.Map) Then
290             Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
292             Call CloseSocket(UserIndex)
                  Exit Sub
              End If
        
              'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
              'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martin Sotuyo Dodero (Maraxus)
294         If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

                  Dim FoundPlace As Boolean

                  Dim esAgua     As Boolean

                  Dim tX         As Long

                  Dim tY         As Long
        
296             FoundPlace = False
298             esAgua = (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0
        
300             For tY = .Pos.Y - 1 To .Pos.Y + 1
302                 For tX = .Pos.X - 1 To .Pos.X + 1

304                     If esAgua Then

                              'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
306                         If LegalPos(.Pos.Map, tX, tY, True, False) Then
308                             FoundPlace = True
                                  Exit For

                              End If

                          Else

                              'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
310                         If LegalPos(.Pos.Map, tX, tY, False, True) Then
312                             FoundPlace = True
                                  Exit For

                              End If

                          End If

314                 Next tX
            
316                 If FoundPlace Then Exit For
318             Next tY
        
320             If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
322                 .Pos.X = tX
324                 .Pos.Y = tY
                  Else

                      'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
326                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                          'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
328                     If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                              'Le avisamos al que estaba comerciando que se tuvo que ir.
330                         If UserList(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
332                             Call FinComerciarUsu(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
334                             Call WriteConsoleMsg(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_WARNING)
                              End If

                              'Lo sacamos.
336                         If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
338                             Call FinComerciarUsu(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
340                             Call WriteErrorMsg(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")
                              End If

                          End If
                
342                     Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)

                      End If

                  End If

              End If
        
              'If in the water, and has a boat, equip it!
344         If .Invent.BarcoObjIndex > 0 And (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Then

                  Dim Barco As ObjData

346             Barco = ObjData(.Invent.BarcoObjIndex)

348             If Barco.Ropaje <> iTraje Then
350                 .Char.Head = 0
                  End If

352             If .flags.Muerto = 0 Then

                      '(Nacho)
354                 If .Faccion.ArmadaReal = 1 Then
356                     If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCiuda
358                     If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCiuda
360                     If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCiuda
362                     If Barco.Ropaje = iTraje Then .Char.Body = iTraje
364                 ElseIf .Faccion.FuerzasCaos = 1 Then

366                     If Barco.Ropaje = iBarca Then .Char.Body = iBarcaPk
368                     If Barco.Ropaje = iGalera Then .Char.Body = iGaleraPk
370                     If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonPk
372                     If Barco.Ropaje = iTraje Then .Char.Body = iTraje
                      Else

374                     If Barco.Ropaje = iBarca Then .Char.Body = iBarca
376                     If Barco.Ropaje = iGalera Then .Char.Body = iGalera
378                     If Barco.Ropaje = iGaleon Then .Char.Body = iGaleon
380                     If Barco.Ropaje = iTraje Then .Char.Body = iTraje

                      End If

                  Else
382                 .Char.Body = iFragataFantasmal
                  End If
            
384             .Char.ShieldAnim = NingunEscudo
386             .Char.WeaponAnim = NingunArma
388             .Char.CascoAnim = NingunCasco
390             .flags.Navegando = 1
            
392             .Char.speeding = Barco.Velocidad
            
394             If Barco.Ropaje = iTraje Then
396                 Call WriteNadarToggle(UserIndex, True)
                  Else
398                 Call WriteNadarToggle(UserIndex, False)

                  End If
              End If

400         Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
402         .flags.NecesitaOxigeno = RequiereOxigeno(.Pos.Map)
        
404         Call WriteHora(UserIndex)
406         Call WriteChangeMap(UserIndex, .Pos.Map) 'Carga el mapa
        
              'If .flags.Privilegios <> PlayerType.user And .flags.Privilegios <> (PlayerType.user Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.user Or PlayerType.RoyalCouncil) And .flags.Privilegios <> (PlayerType.user Or PlayerType.Admin) And .flags.Privilegios <> (PlayerType.user Or PlayerType.Dios) Then
              ' .flags.ChatColor = RGB(2, 161, 38)
              'ElseIf .flags.Privilegios = (PlayerType.user Or PlayerType.RoyalCouncil) Then
              ' .flags.ChatColor = RGB(0, 255, 255)
408         If .flags.Privilegios = PlayerType.Admin Then
410             .flags.ChatColor = RGB(217, 164, 32)
412         ElseIf .flags.Privilegios = PlayerType.Dios Then
414             .flags.ChatColor = RGB(217, 164, 32)
416         ElseIf .flags.Privilegios = PlayerType.SemiDios Then
418             .flags.ChatColor = RGB(2, 161, 38)
420         ElseIf .flags.Privilegios = PlayerType.Consejero Then
422             .flags.ChatColor = RGB(2, 161, 38)
              Else
424             .flags.ChatColor = vbWhite
              End If
        
              ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
              #If ConUpTime Then
426             .LogOnTime = Now
              #End If
        
428         If .flags.Navegando = 0 Then
430             If .flags.Muerto = 0 Then
432                 .Char.speeding = VelocidadNormal
                  Else
434                 .Char.speeding = VelocidadMuerto
                  End If
              End If
        
              'Crea  el personaje del usuario
436         Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, 1)

438         Call WriteUserCharIndexInServer(UserIndex)
        
440         If (.flags.Privilegios And PlayerType.user) = 0 Then
442             Call DoAdminInvisible(UserIndex)
              Else
444             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
              End If
        
446         Call WriteVelocidadToggle(UserIndex)
        
448         Call WriteUpdateUserStats(UserIndex)
        
450         Call WriteUpdateHungerAndThirst(UserIndex)
        
452         Call WriteUpdateDM(UserIndex)
454         Call WriteUpdateRM(UserIndex)
        
456         Call SendMOTD(UserIndex)
        
458         Call SetUserLogged(UserIndex)
        
              'Actualiza el Num de usuarios
460         NumUsers = NumUsers + 1
462         .flags.UserLogged = True
464         .Counters.LastSave = GetTickCount
        
466         MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
468         If .Stats.SkillPts > 0 Then
470             Call WriteSendSkills(UserIndex)
472             Call WriteLevelUp(UserIndex, .Stats.SkillPts)
              End If
        
474         If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
        
476         If NumUsers > RecordUsuarios Then
478             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultáneamente: " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
480             RecordUsuarios = NumUsers
            
482             If Database_Enabled Then
484                 Call SetRecordUsersDatabase(RecordUsuarios)
                  Else
486                 Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(RecordUsuarios))
                  End If
              End If
        
488         Call WriteFYA(UserIndex)
490         Call WriteBindKeys(UserIndex)
        
492         If .NroMascotas > 0 And MapInfo(.Pos.Map).Seguro = 0 And .flags.MascotasGuardadas = 0 Then
                  Dim i As Integer
494             For i = 1 To MAXMASCOTAS
496                 If .MascotasType(i) > 0 Then
498                     .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, False, False)
                    
500                     If .MascotasIndex(i) > 0 Then
502                         Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
504                         Call FollowAmo(.MascotasIndex(i))
                          Else
506                         .MascotasIndex(i) = 0
                          End If
                      End If
508             Next i
              End If
        
510         If .flags.Navegando = 1 Then
512             Call WriteNavigateToggle(UserIndex)
              End If
        
514         If .flags.Montado = 1 Then
516             .Char.speeding = VelocidadMontura
518             Call WriteEquiteToggle(UserIndex)
                  'Debug.Print "Montado:" & .Char.speeding
520             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
              End If
        
522         If .flags.Muerto = 1 Then
524             .Char.speeding = VelocidadMuerto
526             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
              End If
        
528         If .GuildIndex > 0 Then

                  'welcome to the show baby...
530             If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
532                 Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
                  End If

              End If
        
534         tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
    
536         If LenB(tStr) <> 0 Then
538             Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
              End If
        
540         .flags.SolicitudPendienteDe = 0

542         If Lloviendo Then
544             Call WriteRainToggle(UserIndex)
              End If
        
546         If ServidorNublado Then
548             Call WriteNubesToggle(UserIndex)
              End If

550         Call WriteLoggedMessage(UserIndex)
        
552         If .Stats.ELV = 1 Then
554             Call WriteConsoleMsg(UserIndex, "¡Bienvenido a las tierras de AO20! ¡" & .name & " que tengas buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
556         ElseIf .Stats.ELV < 14 Then
558             Call WriteConsoleMsg(UserIndex, "¡Bienvenido de nuevo " & .name & "! Actualmente estas en el nivel " & .Stats.ELV & " en " & DarNameMapa(.Pos.Map) & ", ¡buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)
              End If

560         If Status(UserIndex) = 2 Or Status(UserIndex) = 0 Then
562             Call WriteSafeModeOff(UserIndex)
564             .flags.Seguro = False
              Else
566             .flags.Seguro = True
568             Call WriteSafeModeOn(UserIndex)
              End If
        
              'Call modGuilds.SendGuildNews(UserIndex)
        
570         If .MENSAJEINFORMACION <> vbNullString Then
572             Call WriteConsoleMsg(UserIndex, .MENSAJEINFORMACION, FontTypeNames.FONTTYPE_CENTINELA)
574             .MENSAJEINFORMACION = vbNullString
              End If

576         tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
        
578         If LenB(tStr) <> 0 Then
580             Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
              End If

582         If EventoActivo Then
584             Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)
              End If
        
586         Call WriteContadores(UserIndex)
588         Call WriteOxigeno(UserIndex)

              'Call SendData(UserIndex, UserIndex, PrepareMessageParticleFXToFloor(.Pos.x, .Pos.y, 209, 10))
              'Call SendData(UserIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 209, 40, False))
        
              'Load the user statistics
              'Call Statistics.UserConnected(UserIndex)

590         Call MostrarNumUsers
              'Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageParticleFXToFloor(.Pos.X, .Pos.y, ParticulasIndex.LogeoLevel1, 400))
              'Call SaveUser(UserIndex, CharPath & UCase$(.name) & ".chr")
        
              ' n = FreeFile
              ' Open App.Path & "\logs\numusers.log" For Output As n
              'Print #n, NumUsers
              ' Close #n

          End With
    
          Exit Sub
    
ErrHandler:
592       Call RegistrarError(Err.Number, Err.description, "TCP.ConnectUser", Erl)
594     Call WriteShowMessageBox(UserIndex, "El personaje contiene un error, comuniquese con un miembro del staff.")
596       Call CloseSocket(UserIndex)

End Sub

Sub SendMOTD(ByVal UserIndex As Integer)
        
        On Error GoTo SendMOTD_Err
        

        Dim j As Long

100     For j = 1 To MaxLines
102         Call WriteConsoleMsg(UserIndex, MOTD(j).texto, FontTypeNames.FONTTYPE_EXP)
104     Next j
    
        
        Exit Sub

SendMOTD_Err:
106     Call RegistrarError(Err.Number, Err.description, "TCP.SendMOTD", Erl)
108     Resume Next
        
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
134     Call RegistrarError(Err.Number, Err.description, "TCP.ResetFacciones", Erl)
136     Resume Next
        
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
176         .TiempoDeInmunidad = 0
        End With

        
        Exit Sub

ResetContadores_Err:
178     Call RegistrarError(Err.Number, Err.description, "TCP.ResetContadores", Erl)
180     Resume Next
        
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
114         .Heading = 0
116         .loops = 0
118         .ShieldAnim = 0
120         .WeaponAnim = 0
122         .Arma_Aura = ""
124         .Body_Aura = ""
126         .Head_Aura = ""
128         .Otra_Aura = ""
130         .Anillo_Aura = ""
132         .Escudo_Aura = ""
134         .ParticulaFx = 0
136         .speeding = VelocidadCero

        End With

        
        Exit Sub

ResetCharInfo_Err:
138     Call RegistrarError(Err.Number, Err.description, "TCP.ResetCharInfo", Erl)
140     Resume Next
        
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
            
198         .NroMascotas = 0
        
        End With

        
        Exit Sub

ResetBasicUserInfo_Err:
200     Call RegistrarError(Err.Number, Err.description, "TCP.ResetBasicUserInfo", Erl)
202     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.description, "TCP.ResetGuildInfo", Erl)
114     Resume Next
        
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
206         .Montado = 0
208         .Incinerado = 0
210         .Casado = 0
212         .Pareja = ""
214         .Candidato = 0
216         .UsandoMacro = False
218         .pregunta = 0
            'Ladder
220         .BattleModo = 0

222         .Subastando = False
224         .Paraliza = 0
226         .Envenena = 0
228         .NoPalabrasMagicas = 0
230         .NoMagiaEfeceto = 0
232         .incinera = 0
234         .Estupidiza = 0
236         .GolpeCertero = 0
238         .PendienteDelExperto = 0
240         .CarroMineria = 0
242         .PendienteDelSacrificio = 0
244         .AnilloOcultismo = 0
246         .RegeneracionMana = 0
248         .RegeneracionHP = 0
250         .RegeneracionSta = 0
252         .NecesitaOxigeno = False
254         .LastCrimMatado = vbNullString
256         .LastCiudMatado = vbNullString
        
258         .UserLogged = False
260         .FirstPacket = False
262         .Inmunidad = 0
            
264         .Mimetizado = 0
266         .MascotasGuardadas = 0

268         .EnConsulta = False
        End With

        
        Exit Sub

ResetUserFlags_Err:
270     Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserFlags", Erl)
272     Resume Next
        
End Sub

Sub ResetAccionesPendientes(ByVal UserIndex As Integer)
        
        On Error GoTo ResetAccionesPendientes_Err
        

        '*************************************************
        '*************************************************
100     With UserList(UserIndex).Accion
102         .AccionPendiente = False
104         .HechizoPendiente = 0
106         .RunaObj = 0
108         .Particula = 0
110         .TipoAccion = 0
112         .ObjSlot = 0

        End With

        
        Exit Sub

ResetAccionesPendientes_Err:
114     Call RegistrarError(Err.Number, Err.description, "TCP.ResetAccionesPendientes", Erl)
116     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.description, "TCP.ResetDonadorFlag", Erl)
110     Resume Next
        
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
106     Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSpells", Erl)
108     Resume Next
        
End Sub

Sub ResetUserSkills(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserSkills_Err
        

        Dim LoopC As Long

100     For LoopC = 1 To NUMSKILLS
102         UserList(UserIndex).Stats.UserSkills(LoopC) = 0
104     Next LoopC

        
        Exit Sub

ResetUserSkills_Err:
106     Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSkills", Erl)
108     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserBanco", Erl)
114     Resume Next
        
End Sub

Sub ResetUserKeys(ByVal UserIndex As Integer)
        
        On Error GoTo ResetUserKeys_Err
    
        
100     With UserList(UserIndex)
            Dim i As Integer
        
102         For i = 1 To MAXKEYS
104             .Keys(i) = 0
            Next
        End With
        
        Exit Sub

ResetUserKeys_Err:
106     Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserKeys", Erl)

        
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
108     Call RegistrarError(Err.Number, Err.description, "TCP.LimpiarComercioSeguro", Erl)
110     Resume Next
        
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
160     Call ResetUserKeys(UserIndex)

162     With UserList(UserIndex).ComUsu
164         .Acepto = False
166         .cant = 0
168         .DestNick = vbNullString
170         .DestUsu = 0
172         .Objeto = 0

        End With

        
        Exit Sub

ResetUserSlot_Err:
174     Call RegistrarError(Err.Number, Err.description, "TCP.ResetUserSlot", Erl)
176     Resume Next
        
End Sub

Sub CloseUser(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
        Dim errordesc As String
        Dim Map As Integer
        Dim aN  As Integer
        Dim i   As Integer
    
100     Map = UserList(UserIndex).Pos.Map
    
102     errordesc = "ERROR AL SETEAR NPC"
    
104     aN = UserList(UserIndex).flags.AtacadoPorNpc

106     If aN > 0 Then
108         Npclist(aN).Movement = Npclist(aN).flags.OldMovement
110         Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
112         Npclist(aN).flags.AttackedBy = vbNullString

        End If

114     aN = UserList(UserIndex).flags.NPCAtacado

116     If aN > 0 Then
118         If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
120             Npclist(aN).flags.AttackedFirstBy = vbNullString

            End If

        End If

122     UserList(UserIndex).flags.AtacadoPorNpc = 0
124     UserList(UserIndex).flags.NPCAtacado = 0
    
126     errordesc = "ERROR AL DESMONTAR"

128     If UserList(UserIndex).flags.Montado > 0 Then
130         Call DoMontar(UserIndex, ObjData(UserList(UserIndex).Invent.MonturaObjIndex), UserList(UserIndex).Invent.MonturaSlot)
        End If
    
132     errordesc = "ERROR AL SACAR MIMETISMO"
134     If UserList(UserIndex).flags.Mimetizado = 1 Then
136         UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
138         UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
140         UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
142         UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
144         UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
146         UserList(UserIndex).Counters.Mimetismo = 0
148         UserList(UserIndex).flags.Mimetizado = 0
        End If
    
150     errordesc = "ERROR AL ENVIAR PARTICULA"
    
152     UserList(UserIndex).Char.FX = 0
154     UserList(UserIndex).Char.loops = 0
156     UserList(UserIndex).Char.ParticulaFx = 0
158     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 0, 0, True))
160     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    
162     UserList(UserIndex).flags.UserLogged = False
164     UserList(UserIndex).Counters.Saliendo = False
    
166     errordesc = "ERROR AL ENVIAR INVI"
    
        'Le devolvemos el body y head originales
168     If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)
    
170     errordesc = "ERROR AL CANCELAR SUBASTA"

172     If UserList(UserIndex).flags.Subastando = True Then
174         Call CancelarSubasta

        End If
    
176     errordesc = "ERROR AL BORRAR INDEX DE TORNEO"

178     If UserList(UserIndex).flags.EnTorneo = True Then
180         Call BorrarIndexInTorneo(UserIndex)
182         UserList(UserIndex).flags.EnTorneo = False

        End If
    
        'Save statistics
        'Call Statistics.UserDisconnected(UserIndex)
    
        ' Grabamos el personaje del usuario
    
184     errordesc = "ERROR AL GRABAR PJ"
    
186     If UserList(UserIndex).flags.BattleModo = 0 Then
188         Call SaveUser(UserIndex, True)
        Else
            'Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)
190         Call SaveBattlePoints(UserIndex)

        End If

192     errordesc = "ERROR AL DESCONTAR USER DE MAPA"

194     If MapInfo(Map).NumUsers > 0 Then
196         Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))

        End If

198     errordesc = "ERROR AL ERASEUSERCHAR"
    
        'Borrar el personaje
200     Call EraseUserChar(UserIndex, True)
    
202     errordesc = "ERROR AL BORRAR MASCOTAS"
    
        'Borrar mascotas
204     For i = 1 To MAXMASCOTAS
206         If UserList(UserIndex).MascotasIndex(i) > 0 Then
208             If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
                    Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
            End If
210     Next i
    
212     errordesc = "ERROR Update Map Users"
    
        'Update Map Users
214     MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
    
216     If MapInfo(Map).NumUsers < 0 Then MapInfo(Map).NumUsers = 0

        ' Si el usuario habia dejado un msg en la gm's queue lo borramos
        'If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)
    
218     errordesc = "ERROR AL RESETEAR FLAGS Name:" & UserList(UserIndex).name & " cuenta:" & UserList(UserIndex).Cuenta
    
        'Reseteo los estados del juagador, fuerza el cierre del cliente.
220     Call ResetUserFlags(UserIndex)
    
222     errordesc = "ERROR AL RESETSLOT Name:" & UserList(UserIndex).name & " cuenta:" & UserList(UserIndex).Cuenta
    
224     Call ResetUserSlot(UserIndex)
    
        Exit Sub
    
ErrHandler:
226     Call LogError("Error en CloseUser. Número " & Err.Number & ". Descripción: " & Err.description & ". Detalle:" & errordesc)

228     Resume Next ' TODO: Provisional hasta solucionar bugs graves

End Sub

Sub ReloadSokcet()

        On Error GoTo ErrHandler

        #If UsarQueSocket = 1 Then

100         Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
102         If NumUsers <= 0 Then
104             Call WSApiReiniciarSockets
            Else

                '       Call apiclosesocket(SockListen)
                '       SockListen = ListenForConnect(Puerto, hWndMsg, "")
            End If

        #ElseIf UsarQueSocket = 0 Then

106         frmMain.Socket1.Cleanup
108         Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
        #ElseIf UsarQueSocket = 2 Then

        #End If

        Exit Sub
ErrHandler:
110     Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

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
110     Call RegistrarError(Err.Number, Err.description, "TCP.EcharPjsNoPrivilegiados", Erl)
112     Resume Next
        
End Sub

Function ValidarCabeza(ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal Head As Integer) As Boolean

100     Select Case UserSexo
    
            Case eGenero.Hombre
        
102             Select Case UserRaza
                
                    Case eRaza.Humano
104                     ValidarCabeza = Head >= 1 And Head <= 41
                    
106                 Case eRaza.Elfo
108                     ValidarCabeza = Head >= 101 And Head <= 132
                    
110                 Case eRaza.Drow
112                     ValidarCabeza = Head >= 200 And Head <= 229
                    
114                 Case eRaza.Enano
116                     ValidarCabeza = Head >= 300 And Head <= 329
                    
118                 Case eRaza.Gnomo
120                     ValidarCabeza = Head >= 400 And Head <= 429
                    
122                 Case eRaza.Orco
124                     ValidarCabeza = Head >= 500 And Head <= 529
                
                End Select
        
126         Case eGenero.Mujer
        
128             Select Case UserRaza
                
                    Case eRaza.Humano
130                     ValidarCabeza = Head >= 50 And Head <= 80
                    
132                 Case eRaza.Elfo
134                     ValidarCabeza = Head >= 150 And Head <= 179
                    
136                 Case eRaza.Drow
138                     ValidarCabeza = Head >= 250 And Head <= 279
                    
140                 Case eRaza.Enano
142                     ValidarCabeza = Head >= 350 And Head <= 379
                    
144                 Case eRaza.Gnomo
146                     ValidarCabeza = Head >= 450 And Head <= 479
                    
148                 Case eRaza.Orco
150                     ValidarCabeza = Head >= 550 And Head <= 579
                
                End Select
    
        End Select

End Function

Function ValidarNombre(nombre As String) As Boolean
    
100     If Len(nombre) < 1 Or Len(nombre) > 18 Then Exit Function
    
        Dim temp As String
102     temp = UCase$(nombre)
    
        Dim i As Long, Char As Integer, LastChar As Integer
104     For i = 1 To Len(temp)
106         Char = Asc(mid$(temp, i, 1))
        
108         If (Char < 65 Or Char > 90) And Char <> 32 Then
                Exit Function
        
110         ElseIf Char = 32 And LastChar = 32 Then
                Exit Function
            End If
        
112         LastChar = Char
        Next

114     If Asc(mid$(temp, 1, 1)) = 32 Or Asc(mid$(temp, Len(temp), 1)) = 32 Then
            Exit Function
        End If
    
116     ValidarNombre = True

End Function
